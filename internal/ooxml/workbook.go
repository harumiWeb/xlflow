package ooxml

import (
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"sort"
	"strconv"
	"strings"
)

type Workbook struct {
	Sheets       []Sheet
	DefinedNames []DefinedName
}

type Sheet struct {
	Index   int
	Name    string
	SheetID string
	RelID   string
	Path    string
}

type DefinedName struct {
	Name         string
	LocalSheetID *int
	RefersTo     string
}

type relationship struct {
	ID     string
	Target string
	Type   string
}

func (p *Package) ReadWorkbook() (Workbook, error) {
	workbook, err := p.readWorkbookXML()
	if err != nil {
		return Workbook{}, err
	}
	rels, err := p.readWorkbookRelationships()
	if err != nil {
		return Workbook{}, err
	}
	for i := range workbook.Sheets {
		target := rels[workbook.Sheets[i].RelID]
		if target == "" {
			return Workbook{}, fmt.Errorf("worksheet relationship %q not found", workbook.Sheets[i].RelID)
		}
		workbook.Sheets[i].Path = resolveWorkbookRelationshipTarget(target)
	}
	sortDefinedNames(workbook.DefinedNames, workbook.Sheets)
	return workbook, nil
}

func (p *Package) readWorkbookXML() (workbook Workbook, err error) {
	rc, err := p.openPart("xl/workbook.xml")
	if err != nil {
		return Workbook{}, err
	}
	defer func() {
		if closeErr := rc.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	decoder := xml.NewDecoder(rc)
	inDefinedName := false
	var current DefinedName
	var text strings.Builder
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return Workbook{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "sheet":
				sheet := Sheet{Index: len(workbook.Sheets) + 1}
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "name":
						sheet.Name = attr.Value
					case "sheetId":
						sheet.SheetID = attr.Value
					case "id":
						sheet.RelID = attr.Value
					}
				}
				workbook.Sheets = append(workbook.Sheets, sheet)
			case "definedName":
				inDefinedName = true
				current = DefinedName{}
				text.Reset()
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "name":
						current.Name = attr.Value
					case "localSheetId":
						if id, err := strconv.Atoi(attr.Value); err == nil {
							current.LocalSheetID = &id
						}
					}
				}
			}
		case xml.CharData:
			if inDefinedName {
				text.Write([]byte(t))
			}
		case xml.EndElement:
			if t.Name.Local == "definedName" && inDefinedName {
				current.RefersTo = ensureFormulaPrefix(strings.TrimSpace(text.String()))
				workbook.DefinedNames = append(workbook.DefinedNames, current)
				inDefinedName = false
			}
		}
	}
	return workbook, nil
}

func (p *Package) readWorkbookRelationships() (result map[string]string, err error) {
	rc, err := p.openPart("xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	defer func() {
		if closeErr := rc.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()
	decoder := xml.NewDecoder(rc)
	result = map[string]string{}
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		var rel relationship
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "Id":
				rel.ID = attr.Value
			case "Target":
				rel.Target = attr.Value
			case "Type":
				rel.Type = attr.Value
			}
		}
		if rel.ID != "" && rel.Target != "" {
			result[rel.ID] = rel.Target
		}
	}
	return result, nil
}

func resolveWorkbookRelationshipTarget(target string) string {
	target = strings.ReplaceAll(target, "\\", "/")
	if strings.HasPrefix(target, "/") {
		return cleanPartName(target)
	}
	return cleanPartName(path.Join("xl", target))
}

func ensureFormulaPrefix(value string) string {
	if value == "" || strings.HasPrefix(value, "=") {
		return value
	}
	return "=" + value
}

func sortDefinedNames(names []DefinedName, sheets []Sheet) {
	sort.SliceStable(names, func(i, j int) bool {
		leftScope := definedNameScopeOrder(names[i], sheets)
		rightScope := definedNameScopeOrder(names[j], sheets)
		if leftScope != rightScope {
			return leftScope < rightScope
		}
		return strings.ToLower(names[i].Name) < strings.ToLower(names[j].Name)
	})
}

func definedNameScopeOrder(name DefinedName, sheets []Sheet) int {
	if name.LocalSheetID == nil {
		return -1
	}
	return *name.LocalSheetID
}

func DefinedNameScope(name DefinedName, sheets []Sheet) string {
	if name.LocalSheetID == nil {
		return "workbook"
	}
	idx := *name.LocalSheetID
	if idx >= 0 && idx < len(sheets) {
		return sheets[idx].Name
	}
	return "sheet"
}
