package formulas

import "testing"

func TestBuildRegionsGroupsContiguousSameColumnFormulas(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "G2", Row: 2, Col: 7, Kind: "normal", Formula: "=E2*F2"},
		{Cell: "G3", Row: 3, Col: 7, Kind: "normal", Formula: "=E3*F3"},
		{Cell: "G4", Row: 4, Col: 7, Kind: "normal", Formula: "=E4*F4"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d", len(regions))
	}
	if regions[0].Range != "G2:G4" || regions[0].FormulaR1C1 != "=RC[-2]*RC[-1]" || regions[0].Count != 3 {
		t.Fatalf("region = %#v", regions[0])
	}
}

func TestBuildRegionsMarksOneCellDeviationAsOutlier(t *testing.T) {
	cells := []FormulaCell{
		{Cell: "G2", Row: 2, Col: 7, Kind: "normal", Formula: "=E2*F2"},
		{Cell: "G3", Row: 3, Col: 7, Kind: "normal", Formula: "=E3*F3"},
		{Cell: "G4", Row: 4, Col: 7, Kind: "normal", Formula: `=IF(E4="",0,E4*F4)`},
		{Cell: "G5", Row: 5, Col: 7, Kind: "normal", Formula: "=E5*F5"},
		{Cell: "G6", Row: 6, Col: 7, Kind: "normal", Formula: "=E6*F6"},
	}
	regions := BuildRegions(cells)
	if len(regions) != 3 {
		t.Fatalf("region count = %d: %#v", len(regions), regions)
	}
	if regions[1].Range != "G4" || !contains(regions[1].Features, "outlier") {
		t.Fatalf("outlier region = %#v", regions[1])
	}
}

func TestBuildRegionsSharedFormulaUsesAnchorRef(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "C2", Row: 2, Col: 3, Kind: "shared", SharedIndex: "0", SharedRef: "C2:C4", Formula: "=A2*B2"},
		{Cell: "C3", Row: 3, Col: 3, Kind: "shared", SharedIndex: "0"},
		{Cell: "C4", Row: 4, Col: 3, Kind: "shared", SharedIndex: "0"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d: %#v", len(regions), regions)
	}
	if regions[0].Range != "C2:C4" || regions[0].Count != 3 {
		t.Fatalf("shared region = %#v", regions[0])
	}
	if len(regions[0].StorageKinds) != 1 || regions[0].StorageKinds[0] != "shared" {
		t.Fatalf("storage kinds = %#v", regions[0].StorageKinds)
	}
	if regions[0].FormulaR1C1 != "=RC[-2]*RC[-1]" {
		t.Fatalf("formula_r1c1 = %q", regions[0].FormulaR1C1)
	}
}

func TestBuildRegionsCoalescesAdjacentSharedFormulaStorageGroups(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "D2", Row: 2, Col: 4, Kind: "shared", SharedIndex: "0", SharedRef: "D2:D65", Formula: "=B2*C2"},
		{Cell: "D66", Row: 66, Col: 4, Kind: "shared", SharedIndex: "1", SharedRef: "D66:D101", Formula: "=B66*C66"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d: %#v", len(regions), regions)
	}
	if regions[0].Range != "D2:D101" || regions[0].FormulaR1C1 != "=RC[-2]*RC[-1]" || regions[0].Count != 100 {
		t.Fatalf("coalesced region = %#v", regions[0])
	}
	if regions[0].StorageGroupCount != 2 {
		t.Fatalf("storage_group_count = %d", regions[0].StorageGroupCount)
	}
	if len(regions[0].StorageKinds) != 1 || regions[0].StorageKinds[0] != "shared" {
		t.Fatalf("storage_kinds = %#v", regions[0].StorageKinds)
	}
	assertNoOverlappingRegions(t, regions)
}

func TestBuildRegionsExplicitFormulaOverridesSharedExpansionForOutlier(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "D2", Row: 2, Col: 4, Kind: "shared", SharedIndex: "0", SharedRef: "D2:D1001", Formula: "=B2*C2"},
		{Cell: "D500", Row: 500, Col: 4, Kind: "shared", SharedIndex: "1", SharedRef: "D500", Formula: `=IF(B500="",0,B500*C500)`},
	})
	if len(regions) != 3 {
		t.Fatalf("region count = %d: %#v", len(regions), regions)
	}
	if regions[0].Range != "D2:D499" || regions[0].Count != 498 {
		t.Fatalf("first region = %#v", regions[0])
	}
	if regions[1].Range != "D500" || regions[1].FormulaR1C1 != `=IF(RC[-2]="",0,RC[-2]*RC[-1])` || !contains(regions[1].Features, "outlier") {
		t.Fatalf("outlier region = %#v", regions[1])
	}
	if regions[2].Range != "D501:D1001" || regions[2].Count != 501 {
		t.Fatalf("third region = %#v", regions[2])
	}
	assertNoOverlappingRegions(t, regions)
	assertRegionCellCount(t, regions, 1000)
}

func TestBuildRegionsUnsupportedFormulaPreservesRaw(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "D10", Row: 10, Col: 4, Kind: "normal", Formula: "=Table1[Amount]"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d", len(regions))
	}
	if regions[0].ParseStatus != "partial" || regions[0].Formula != "=Table1[Amount]" || regions[0].FormulaR1C1 != "" {
		t.Fatalf("region = %#v", regions[0])
	}
	if !contains(regions[0].Features, "structured_reference") {
		t.Fatalf("features = %#v", regions[0].Features)
	}
}

func TestBuildRegionsAddsFormulaIntelligenceIndex(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "E2", Row: 2, Col: 5, Kind: "normal", Formula: "=IFERROR(D2*Config!$B$2,0)"},
		{Cell: "E3", Row: 3, Col: 5, Kind: "normal", Formula: "=IFERROR(D3*Config!$B$2,0)"},
		{Cell: "E4", Row: 4, Col: 5, Kind: "normal", Formula: "=IFERROR(D4*Config!$B$2,0)"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d: %#v", len(regions), regions)
	}
	assertStrings(t, regions[0].Refs, []string{"Config!$B$2", "D2:D4"})
	assertStrings(t, regions[0].DependsOnSheets, []string{"Config"})
	assertStrings(t, regions[0].Functions, []string{"IFERROR"})
}

func TestBuildRegionsMissingSharedAnchorFailsSoft(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "C3", Row: 3, Col: 3, Kind: "shared", SharedIndex: "0"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d", len(regions))
	}
	if regions[0].ParseStatus != "failed" || !contains(regions[0].Features, "shared_formula_missing_anchor") {
		t.Fatalf("region = %#v", regions[0])
	}
}

func TestBuildRegionsMalformedSharedAnchorFailsSoft(t *testing.T) {
	regions := BuildRegions([]FormulaCell{
		{Cell: "C2", Row: 2, Col: 3, Kind: "shared", SharedIndex: "0", Formula: "=A2*B2"},
	})
	if len(regions) != 1 {
		t.Fatalf("region count = %d", len(regions))
	}
	if regions[0].ParseStatus != "failed" || regions[0].Formula != "=A2*B2" {
		t.Fatalf("region = %#v", regions[0])
	}
	if !contains(regions[0].Features, "shared_formula_malformed_ref") {
		t.Fatalf("features = %#v", regions[0].Features)
	}
}

func contains(values []string, value string) bool {
	for _, v := range values {
		if v == value {
			return true
		}
	}
	return false
}

func assertStrings(t *testing.T, got, want []string) {
	t.Helper()
	if len(got) != len(want) {
		t.Fatalf("strings = %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("strings = %#v, want %#v", got, want)
		}
	}
}

func assertNoOverlappingRegions(t *testing.T, regions []FormulaRegion) {
	t.Helper()
	seen := map[string]string{}
	for _, region := range regions {
		for _, cell := range expandRegionCells(t, region.Range) {
			if previous, ok := seen[cell]; ok {
				t.Fatalf("cell %s appears in multiple regions: %s and %s", cell, previous, region.Range)
			}
			seen[cell] = region.Range
		}
	}
}

func assertRegionCellCount(t *testing.T, regions []FormulaRegion, want int) {
	t.Helper()
	got := 0
	for _, region := range regions {
		got += len(expandRegionCells(t, region.Range))
	}
	if got != want {
		t.Fatalf("expanded region cell count = %d, want %d", got, want)
	}
}

func expandRegionCells(t *testing.T, rangeText string) []string {
	t.Helper()
	start, end, _, ok := rangeInfo(rangeText)
	if !ok {
		t.Fatalf("invalid range %q", rangeText)
	}
	var cells []string
	for row := start.row; row <= end.row; row++ {
		for col := start.col; col <= end.col; col++ {
			cell := renderRange(col, row, col, row)
			cells = append(cells, cell)
		}
	}
	return cells
}
