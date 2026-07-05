package formulas

import "github.com/harumiWeb/xlflow/internal/formula"

type Manifest struct {
	Version  int             `json:"version"`
	Workbook string          `json:"workbook"`
	Sheets   []SheetManifest `json:"sheets"`
}

type SheetManifest struct {
	Index              int    `json:"index"`
	Name               string `json:"name"`
	SheetID            string `json:"sheet_id"`
	Path               string `json:"path"`
	FormulaRegionCount int    `json:"formula_region_count"`
}

type DefinedName struct {
	Name     string `json:"name"`
	Scope    string `json:"scope"`
	RefersTo string `json:"refers_to"`
	Kind     string `json:"kind"`
}

type FormulaCell struct {
	Cell        string
	Row         int
	Col         int
	Kind        string
	SharedIndex string
	SharedRef   string
	Formula     string
}

type FormulaRegion struct {
	Range             string   `json:"range"`
	Kind              string   `json:"kind"`
	FormulaR1C1       string   `json:"formula_r1c1,omitempty"`
	Formula           string   `json:"formula,omitempty"`
	ExampleCell       string   `json:"example_cell,omitempty"`
	ExampleFormula    string   `json:"example_formula,omitempty"`
	Count             int      `json:"count"`
	ParseStatus       string   `json:"parse_status"`
	Features          []string `json:"features,omitempty"`
	StorageKinds      []string `json:"storage_kinds,omitempty"`
	StorageGroupCount int      `json:"storage_group_count,omitempty"`

	startRow      int
	endRow        int
	col           int
	key           regionKey
	storageGroups int
}

type regionKey struct {
	Kind        string
	FormulaR1C1 string
	Formula     string
	ParseStatus string
	Features    string
}

type normalizeSummary struct {
	status      formula.ParseStatus
	formulaR1C1 string
	rawFormula  string
	features    []string
}

type Result struct {
	Manifest           Manifest
	Names              []DefinedName
	OutputDir          string
	ManifestPath       string
	FormulaRegionCount int
}
