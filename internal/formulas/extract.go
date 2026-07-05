package formulas

import (
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/ooxml"
	"github.com/xuri/excelize/v2"
)

func Extract(workbookPath string) (manifest Manifest, names []DefinedName, regionsByPath map[string][]FormulaRegion, err error) {
	pkg, err := ooxml.Open(workbookPath)
	if err != nil {
		return Manifest{}, nil, nil, err
	}
	defer func() {
		if closeErr := pkg.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	wb, err := pkg.ReadWorkbook()
	if err != nil {
		return Manifest{}, nil, nil, err
	}
	manifest = Manifest{
		Version:  1,
		Workbook: filepath.Base(workbookPath),
		Sheets:   make([]SheetManifest, 0, len(wb.Sheets)),
	}
	names = make([]DefinedName, 0, len(wb.DefinedNames))
	for _, name := range wb.DefinedNames {
		names = append(names, DefinedName{
			Name:     name.Name,
			Scope:    ooxml.DefinedNameScope(name, wb.Sheets),
			RefersTo: name.RefersTo,
			Kind:     "formula",
		})
	}

	usedPaths := map[string]int{}
	regionsByPath = map[string][]FormulaRegion{}
	for _, sheet := range wb.Sheets {
		formulasXML, err := pkg.ReadWorksheetFormulas(sheet.Path)
		if err != nil {
			return Manifest{}, nil, nil, err
		}
		cells := make([]FormulaCell, 0, len(formulasXML))
		for _, f := range formulasXML {
			row, col, err := cellCoordinates(f.Cell)
			if err != nil {
				cells = append(cells, FormulaCell{
					Cell:        f.Cell,
					Kind:        formulaKind(f.Type),
					SharedIndex: f.SharedIndex,
					SharedRef:   f.Ref,
					Formula:     f.Text,
				})
				continue
			}
			cells = append(cells, FormulaCell{
				Cell:        f.Cell,
				Row:         row,
				Col:         col,
				Kind:        formulaKind(f.Type),
				SharedIndex: f.SharedIndex,
				SharedRef:   f.Ref,
				Formula:     f.Text,
			})
		}
		regions := BuildRegions(cells)
		relPath := sheetOutputPath(sheet.Index, sheet.Name, usedPaths)
		regionsByPath[relPath] = regions
		manifest.Sheets = append(manifest.Sheets, SheetManifest{
			Index:              sheet.Index,
			Name:               sheet.Name,
			SheetID:            sheet.SheetID,
			Path:               relPath,
			FormulaRegionCount: len(regions),
		})
	}
	return manifest, names, regionsByPath, nil
}

func formulaKind(value string) string {
	if strings.EqualFold(value, "shared") {
		return "shared"
	}
	return "normal"
}

func cellCoordinates(cell string) (row, col int, err error) {
	col, row, err = excelize.CellNameToCoordinates(cell)
	return row, col, err
}

func sheetOutputPath(index int, name string, used map[string]int) string {
	base := fmt.Sprintf("%03d-%s", index, sanitizeSheetName(name))
	used[base]++
	if used[base] > 1 {
		base = fmt.Sprintf("%s-%d", base, used[base])
	}
	return filepath.ToSlash(filepath.Join("sheets", base+".regions.jsonl"))
}

func sanitizeSheetName(name string) string {
	name = strings.TrimSpace(name)
	var b strings.Builder
	lastDash := false
	for _, r := range name {
		valid := r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r >= '0' && r <= '9'
		if valid {
			b.WriteRune(r)
			lastDash = false
			continue
		}
		if !lastDash && b.Len() > 0 {
			b.WriteByte('-')
			lastDash = true
		}
	}
	result := strings.Trim(b.String(), "-")
	if result == "" {
		return "Sheet"
	}
	return result
}

func sortCells(cells []FormulaCell) {
	sort.SliceStable(cells, func(i, j int) bool {
		if cells[i].Row != cells[j].Row {
			return cells[i].Row < cells[j].Row
		}
		if cells[i].Col != cells[j].Col {
			return cells[i].Col < cells[j].Col
		}
		return cells[i].Cell < cells[j].Cell
	})
}

func sortCellsByColumn(cells []FormulaCell) {
	sort.SliceStable(cells, func(i, j int) bool {
		if cells[i].Col != cells[j].Col {
			return cells[i].Col < cells[j].Col
		}
		if cells[i].Row != cells[j].Row {
			return cells[i].Row < cells[j].Row
		}
		return cells[i].Cell < cells[j].Cell
	})
}
