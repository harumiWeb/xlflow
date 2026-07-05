package formulas

import (
	"sort"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/formula"
	"github.com/xuri/excelize/v2"
)

func finalizeRegionIntelligence(regions []FormulaRegion) {
	for i := range regions {
		if regions[i].ExampleFormula == "" || regions[i].ExampleCell == "" {
			continue
		}
		result := formula.NormalizeA1ToR1C1Pattern(regions[i].ExampleFormula, formula.NormalizeOptions{
			BaseCell: formula.CellRef{Row: regions[i].startRow, Col: regions[i].col},
		})
		regions[i].Refs = regionRefs(regions[i], result.References)
		regions[i].DependsOnSheets = sheetsFromRefs(result.References)
		regions[i].Functions = formulaFunctions(regions[i].ExampleFormula)
	}
}

func regionRefs(region FormulaRegion, refs []formula.Reference) []string {
	if len(refs) == 0 {
		return nil
	}
	values := make([]string, 0, len(refs))
	for _, ref := range refs {
		if rendered, ok := renderRegionRef(region, ref); ok {
			values = append(values, rendered)
		}
	}
	return uniqueSorted(values)
}

func renderRegionRef(region FormulaRegion, ref formula.Reference) (string, bool) {
	prefix := ""
	if ref.Sheet != "" {
		prefix = ref.Sheet + "!"
	}
	switch ref.Kind {
	case formula.ReferenceKindCell:
		top := resolveEndpointForRegion(ref.Start, region, region.startRow)
		bottom := resolveEndpointForRegion(ref.Start, region, region.endRow)
		return prefix + renderResolvedRange(top, bottom, ref.Start, ref.Start), true
	case formula.ReferenceKindRange:
		if ref.End == nil {
			return "", false
		}
		topStart := resolveEndpointForRegion(ref.Start, region, region.startRow)
		bottomEnd := resolveEndpointForRegion(*ref.End, region, region.endRow)
		return prefix + renderResolvedRange(topStart, bottomEnd, ref.Start, *ref.End), true
	default:
		if ref.Raw == "" {
			return "", false
		}
		return prefix + ref.Raw, true
	}
}

func resolveEndpointForRegion(endpoint formula.RefEndpoint, region FormulaRegion, row int) cellPoint {
	resolved := cellPoint{row: endpoint.Row, col: endpoint.Col}
	if !endpoint.RowAbs && endpoint.Row > 0 {
		resolved.row = row + (endpoint.Row - region.startRow)
	}
	// Regions are currently grouped vertically in one column, so only row-relative
	// references expand across the region. Column-relative expansion needs a
	// target-column input if rectangular regions are introduced later.
	return resolved
}

func renderResolvedRange(start, end cellPoint, startFormat, endFormat formula.RefEndpoint) string {
	if start.row == end.row && start.col == end.col {
		return renderResolvedCell(start, startFormat)
	}
	return renderResolvedCell(start, startFormat) + ":" + renderResolvedCell(end, endFormat)
}

func renderResolvedCell(cell cellPoint, format formula.RefEndpoint) string {
	col, _ := excelize.ColumnNumberToName(cell.col)
	if format.ColAbs {
		col = "$" + col
	}
	row := strconv.Itoa(cell.row)
	if format.RowAbs {
		row = "$" + row
	}
	return col + row
}

func sheetsFromRefs(refs []formula.Reference) []string {
	values := make([]string, 0, len(refs))
	for _, ref := range refs {
		if ref.Sheet != "" {
			values = append(values, ref.Sheet)
		}
	}
	return uniqueSorted(values)
}

func formulaFunctions(value string) []string {
	tokens, _ := formula.Lex(value)
	values := []string{}
	for i, token := range tokens {
		if token.Kind != formula.TokenIdentifier {
			continue
		}
		next := nextNonWhitespace(tokens, i+1)
		if next < 0 || tokens[next].Text != "(" {
			continue
		}
		prev := previousNonWhitespace(tokens, i-1)
		if prev >= 0 && tokens[prev].Text == "!" {
			continue
		}
		values = append(values, strings.ToUpper(token.Text))
	}
	return uniqueSorted(values)
}

func nextNonWhitespace(tokens []formula.Token, pos int) int {
	for i := pos; i < len(tokens); i++ {
		if tokens[i].Kind != formula.TokenWhitespace {
			return i
		}
	}
	return -1
}

func previousNonWhitespace(tokens []formula.Token, pos int) int {
	for i := pos; i >= 0; i-- {
		if tokens[i].Kind != formula.TokenWhitespace {
			return i
		}
	}
	return -1
}

func uniqueSorted(values []string) []string {
	if len(values) == 0 {
		return nil
	}
	seen := map[string]bool{}
	result := make([]string, 0, len(values))
	for _, value := range values {
		if value == "" || seen[value] {
			continue
		}
		seen[value] = true
		result = append(result, value)
	}
	sort.Strings(result)
	return result
}
