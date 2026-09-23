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

// WorksheetCodeNameMapping associates a worksheet's visible name with its
// VBA document-module CodeName.
type WorksheetCodeNameMapping struct {
	Name     string
	CodeName string
	Path     string
}

// WorksheetCodeNameIssueKind identifies a catalog entry that cannot be used
// for a unique visible-name to CodeName lookup.
type WorksheetCodeNameIssueKind string

const (
	WorksheetCodeNameMissingRelationship   WorksheetCodeNameIssueKind = "missing_relationship"
	WorksheetCodeNameMalformedRelationship WorksheetCodeNameIssueKind = "malformed_relationship"
	WorksheetCodeNameMissingCodeName       WorksheetCodeNameIssueKind = "missing_code_name"
	WorksheetCodeNameMalformedSheetPart    WorksheetCodeNameIssueKind = "malformed_sheet_part"
	WorksheetCodeNameAmbiguousVisibleName  WorksheetCodeNameIssueKind = "ambiguous_visible_name"
	WorksheetCodeNameAmbiguousCodeName     WorksheetCodeNameIssueKind = "ambiguous_code_name"
)

// WorksheetCodeNameIssue records a worksheet metadata condition that was
// handled conservatively by the catalog loader.
type WorksheetCodeNameIssue struct {
	Kind   WorksheetCodeNameIssueKind
	Name   string
	Path   string
	Detail string
}

// WorksheetCodeNameCatalog is an immutable, case-insensitive lookup of
// worksheet visible names to document-module CodeNames. Invalid or ambiguous
// entries are retained in Issues and never resolve through Lookup.
type WorksheetCodeNameCatalog struct {
	mappings []WorksheetCodeNameMapping
	byName   map[string]string
	issues   []WorksheetCodeNameIssue
}

// Lookup returns the unique CodeName for a visible worksheet name.
func (c WorksheetCodeNameCatalog) Lookup(name string) (string, bool) {
	codeName, ok := c.byName[strings.ToLower(name)]
	return codeName, ok
}

// Mappings returns the worksheet mappings that had valid individual metadata.
func (c WorksheetCodeNameCatalog) Mappings() []WorksheetCodeNameMapping {
	return append([]WorksheetCodeNameMapping(nil), c.mappings...)
}

// Issues returns metadata conditions that were excluded from Lookup.
func (c WorksheetCodeNameCatalog) Issues() []WorksheetCodeNameIssue {
	return append([]WorksheetCodeNameIssue(nil), c.issues...)
}

type DefinedName struct {
	Name            string
	LocalSheetID    *int
	LocalSheetIDRaw string
	RefersTo        string
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
		rel, ok := rels[workbook.Sheets[i].RelID]
		if !ok || rel.Target == "" {
			return Workbook{}, fmt.Errorf("worksheet relationship %q not found", workbook.Sheets[i].RelID)
		}
		workbook.Sheets[i].Path = resolveWorkbookRelationshipTarget(rel.Target)
	}
	sortDefinedNames(workbook.DefinedNames, workbook.Sheets)
	return workbook, nil
}

// ReadWorksheetCodeNameCatalog reads only worksheet relationships and their
// sheetPr codeName values. Non-worksheet relationships, including chart
// sheets, are excluded from the catalog.
func (p *Package) ReadWorksheetCodeNameCatalog() (WorksheetCodeNameCatalog, error) {
	workbook, err := p.readWorkbookXML()
	if err != nil {
		return WorksheetCodeNameCatalog{}, err
	}
	rels, err := p.readWorkbookRelationships()
	if err != nil {
		return WorksheetCodeNameCatalog{}, err
	}

	catalog := WorksheetCodeNameCatalog{
		mappings: make([]WorksheetCodeNameMapping, 0, len(workbook.Sheets)),
		byName:   map[string]string{},
		issues:   make([]WorksheetCodeNameIssue, 0),
	}
	for _, sheet := range workbook.Sheets {
		if strings.TrimSpace(sheet.Name) == "" {
			catalog.issues = append(catalog.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameMalformedRelationship,
				Detail: "worksheet visible name is empty",
			})
			continue
		}
		rel, ok := rels[sheet.RelID]
		if !ok {
			catalog.issues = append(catalog.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameMissingRelationship,
				Name:   sheet.Name,
				Detail: "workbook.xml relationship is absent",
			})
			continue
		}
		if isChartRelationship(rel.Type) {
			continue
		}
		if !isWorksheetRelationship(rel.Type) || rel.Target == "" {
			catalog.issues = append(catalog.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameMalformedRelationship,
				Name:   sheet.Name,
				Detail: "sheet relationship is not a complete worksheet relationship",
			})
			continue
		}
		partPath := resolveWorkbookRelationshipTarget(rel.Target)
		codeName, err := p.readWorksheetCodeName(partPath)
		if err != nil {
			catalog.issues = append(catalog.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameMalformedSheetPart,
				Name:   sheet.Name,
				Path:   partPath,
				Detail: err.Error(),
			})
			continue
		}
		if codeName == "" {
			catalog.issues = append(catalog.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameMissingCodeName,
				Name:   sheet.Name,
				Path:   partPath,
				Detail: "worksheet sheetPr codeName is absent",
			})
			continue
		}
		catalog.mappings = append(catalog.mappings, WorksheetCodeNameMapping{
			Name:     sheet.Name,
			CodeName: codeName,
			Path:     partPath,
		})
	}
	catalog.buildLookup()
	return catalog, nil
}

const (
	transitionalWorksheetRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	strictWorksheetRelationship       = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
)

func isWorksheetRelationship(relationshipType string) bool {
	return relationshipType == transitionalWorksheetRelationship || relationshipType == strictWorksheetRelationship
}

func isChartRelationship(relationshipType string) bool {
	return relationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" || relationshipType == "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet"
}

func (p *Package) readWorksheetCodeName(partPath string) (codeName string, err error) {
	rc, err := p.openPart(partPath)
	if err != nil {
		return "", err
	}
	defer func() {
		if closeErr := rc.Close(); err == nil && closeErr != nil {
			err = closeErr
		}
	}()

	decoder := xml.NewDecoder(rc)
	rootSeen := false
	codeNameSeen := false
	for {
		tok, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			if !rootSeen {
				return "", fmt.Errorf("worksheet part has no root element")
			}
			return codeName, nil
		}
		if tokenErr != nil {
			return "", tokenErr
		}
		switch token := tok.(type) {
		case xml.StartElement:
			if !rootSeen {
				rootSeen = true
				if token.Name.Local != "worksheet" {
					return "", fmt.Errorf("worksheet relationship points to %q part", token.Name.Local)
				}
				continue
			}
			if token.Name.Local != "sheetPr" {
				if token.Name.Local == "sheetData" {
					if err := decoder.Skip(); err != nil {
						return "", err
					}
				}
				continue
			}
			value := ""
			for _, attr := range token.Attr {
				if attr.Name.Local == "codeName" {
					value = strings.TrimSpace(attr.Value)
				}
			}
			if !codeNameSeen {
				codeName = value
				codeNameSeen = true
				continue
			}
			if value != codeName {
				return "", fmt.Errorf("worksheet contains conflicting sheetPr codeName values")
			}
		}
	}
}

func (c *WorksheetCodeNameCatalog) buildLookup() {
	nameCandidates := map[string][]WorksheetCodeNameMapping{}
	codeNameCandidates := map[string][]WorksheetCodeNameMapping{}
	for _, mapping := range c.mappings {
		nameCandidates[strings.ToLower(mapping.Name)] = append(nameCandidates[strings.ToLower(mapping.Name)], mapping)
		codeNameCandidates[strings.ToLower(mapping.CodeName)] = append(codeNameCandidates[strings.ToLower(mapping.CodeName)], mapping)
	}
	for name, candidates := range nameCandidates {
		if len(candidates) != 1 {
			for _, candidate := range candidates {
				c.issues = append(c.issues, WorksheetCodeNameIssue{
					Kind:   WorksheetCodeNameAmbiguousVisibleName,
					Name:   candidate.Name,
					Path:   candidate.Path,
					Detail: fmt.Sprintf("visible name %q has %d worksheet mappings", candidate.Name, len(candidates)),
				})
			}
			continue
		}
		c.byName[name] = candidates[0].CodeName
	}
	for _, candidates := range codeNameCandidates {
		if len(candidates) <= 1 {
			continue
		}
		for _, candidate := range candidates {
			delete(c.byName, strings.ToLower(candidate.Name))
			c.issues = append(c.issues, WorksheetCodeNameIssue{
				Kind:   WorksheetCodeNameAmbiguousCodeName,
				Name:   candidate.Name,
				Path:   candidate.Path,
				Detail: fmt.Sprintf("CodeName %q is used by %d worksheet mappings", candidate.CodeName, len(candidates)),
			})
		}
	}
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
						current.LocalSheetIDRaw = attr.Value
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

func (p *Package) readWorkbookRelationships() (result map[string]relationship, err error) {
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
	result = map[string]relationship{}
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
		if rel.ID != "" {
			result[rel.ID] = rel
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
		if name.LocalSheetIDRaw != "" {
			return len(sheets)
		}
		return -1
	}
	return *name.LocalSheetID
}

func DefinedNameScope(name DefinedName, sheets []Sheet) string {
	if name.LocalSheetID == nil {
		if name.LocalSheetIDRaw != "" {
			return "sheet"
		}
		return "workbook"
	}
	idx := *name.LocalSheetID
	if idx >= 0 && idx < len(sheets) {
		return sheets[idx].Name
	}
	return "sheet"
}
