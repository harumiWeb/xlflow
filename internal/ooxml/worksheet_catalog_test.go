package ooxml

import "testing"

func TestReadWorksheetCodeNameCatalogUsesWorksheetRelationshipsAndCodeNames(t *testing.T) {
	path := writeZipFixture(t, map[string]string{
		"xl/workbook.xml": `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Report" sheetId="1" r:id="rId1"/>
    <sheet name="Summary" sheetId="2" r:id="rId2"/>
    <sheet name="Chart" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/report.xml"/>
  <Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="worksheets/summary.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/chart.xml"/>
</Relationships>`,
		"xl/worksheets/report.xml":  `<worksheet><sheetPr codeName="SheetReport"/></worksheet>`,
		"xl/worksheets/summary.xml": `<worksheet><sheetPr codeName="SheetSummary"/></worksheet>`,
		"xl/chartsheets/chart.xml":  `<chartsheet><sheetPr codeName="ChartSheet"/></chartsheet>`,
	})
	pkg, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = pkg.Close() }()

	catalog, err := pkg.ReadWorksheetCodeNameCatalog()
	if err != nil {
		t.Fatal(err)
	}
	if got, ok := catalog.Lookup("report"); !ok || got != "SheetReport" {
		t.Fatalf("Report CodeName = %q, %v", got, ok)
	}
	if got, ok := catalog.Lookup("SUMMARY"); !ok || got != "SheetSummary" {
		t.Fatalf("Summary CodeName = %q, %v", got, ok)
	}
	if got, ok := catalog.Lookup("Chart"); ok || got != "" {
		t.Fatalf("chart sheet CodeName = %q, %v; chart sheets must be excluded", got, ok)
	}
	if issues := catalog.Issues(); len(issues) != 0 {
		t.Fatalf("catalog issues = %#v", issues)
	}
}

func TestReadWorksheetCodeNameCatalogReportsMissingAndMalformedEntries(t *testing.T) {
	path := writeZipFixture(t, map[string]string{
		"xl/workbook.xml": `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="MissingRelationship" sheetId="1" r:id="missing"/>
    <sheet name="MissingCodeName" sheetId="2" r:id="rId2"/>
    <sheet name="MalformedPart" sheetId="3" r:id="rId3"/>
    <sheet name="MalformedRelationship" sheetId="4" r:id="rId4"/>
  </sheets>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/missing.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/malformed.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
</Relationships>`,
		"xl/worksheets/missing.xml":   `<worksheet/>`,
		"xl/worksheets/malformed.xml": `<worksheet>`,
	})
	pkg, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = pkg.Close() }()

	catalog, err := pkg.ReadWorksheetCodeNameCatalog()
	if err != nil {
		t.Fatal(err)
	}
	if len(catalog.Mappings()) != 0 {
		t.Fatalf("catalog mappings = %#v", catalog.Mappings())
	}
	want := map[WorksheetCodeNameIssueKind]bool{
		WorksheetCodeNameMissingRelationship:   false,
		WorksheetCodeNameMissingCodeName:       false,
		WorksheetCodeNameMalformedSheetPart:    false,
		WorksheetCodeNameMalformedRelationship: false,
	}
	for _, issue := range catalog.Issues() {
		if _, ok := want[issue.Kind]; ok {
			want[issue.Kind] = true
		}
	}
	for kind, found := range want {
		if !found {
			t.Fatalf("catalog issues = %#v, missing %q", catalog.Issues(), kind)
		}
	}
}

func TestReadWorksheetCodeNameCatalogRejectsAmbiguousNamesAndCodeNames(t *testing.T) {
	path := writeZipFixture(t, map[string]string{
		"xl/workbook.xml": `<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Report" sheetId="1" r:id="rId1"/>
    <sheet name="report" sheetId="2" r:id="rId2"/>
    <sheet name="Other" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/one.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/two.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/three.xml"/>
</Relationships>`,
		"xl/worksheets/one.xml":   `<worksheet><sheetPr codeName="SheetOne"/></worksheet>`,
		"xl/worksheets/two.xml":   `<worksheet><sheetPr codeName="SheetTwo"/></worksheet>`,
		"xl/worksheets/three.xml": `<worksheet><sheetPr codeName="SheetOne"/></worksheet>`,
	})
	pkg, err := Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer func() { _ = pkg.Close() }()

	catalog, err := pkg.ReadWorksheetCodeNameCatalog()
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"Report", "report", "Other"} {
		if got, ok := catalog.Lookup(name); ok || got != "" {
			t.Fatalf("ambiguous %q lookup = %q, %v", name, got, ok)
		}
	}
	var visible, codeName int
	for _, issue := range catalog.Issues() {
		switch issue.Kind {
		case WorksheetCodeNameAmbiguousVisibleName:
			visible++
		case WorksheetCodeNameAmbiguousCodeName:
			codeName++
		}
	}
	if visible != 2 || codeName != 2 {
		t.Fatalf("ambiguous issue counts = visible %d, CodeName %d; issues = %#v", visible, codeName, catalog.Issues())
	}
}

func TestReadWorksheetCodeNameCatalogValidatesSkippedSheetDataAndLaterMetadata(t *testing.T) {
	for _, test := range []struct {
		name      string
		worksheet string
	}{
		{name: "malformed sheet data", worksheet: `<worksheet><sheetPr codeName="SheetReport"/><sheetData><row>`},
		{name: "conflicting later metadata", worksheet: `<worksheet><sheetPr codeName="SheetReport"/><sheetData><row/></sheetData><sheetPr codeName="SheetOther"/></worksheet>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			path := writeZipFixture(t, map[string]string{
				"xl/workbook.xml":            `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`,
				"xl/_rels/workbook.xml.rels": `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/report.xml"/></Relationships>`,
				"xl/worksheets/report.xml":   test.worksheet,
			})
			pkg, err := Open(path)
			if err != nil {
				t.Fatal(err)
			}
			defer func() { _ = pkg.Close() }()

			catalog, err := pkg.ReadWorksheetCodeNameCatalog()
			if err != nil {
				t.Fatal(err)
			}
			if _, ok := catalog.Lookup("Report"); ok {
				t.Fatalf("catalog = %#v, want malformed worksheet mapping excluded", catalog.Mappings())
			}
			issues := catalog.Issues()
			if len(issues) != 1 || issues[0].Kind != WorksheetCodeNameMalformedSheetPart {
				t.Fatalf("issues = %#v, want one malformed sheet part", issues)
			}
		})
	}
}
