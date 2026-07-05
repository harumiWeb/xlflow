package formula

import "testing"

func TestNormalizeRequiredFixtures(t *testing.T) {
	tests := []struct {
		name    string
		base    CellRef
		formula string
		want    string
	}{
		{name: "copied row 2", base: CellRef{Row: 2, Col: 3}, formula: "=A2*B2", want: "=RC[-2]*RC[-1]"},
		{name: "copied row 3", base: CellRef{Row: 3, Col: 3}, formula: "=A3*B3", want: "=RC[-2]*RC[-1]"},
		{name: "absolute", base: CellRef{Row: 2, Col: 3}, formula: "=$A$1", want: "=R1C1"},
		{name: "absolute row", base: CellRef{Row: 2, Col: 3}, formula: "=A$1", want: "=R1C[-2]"},
		{name: "absolute column", base: CellRef{Row: 2, Col: 3}, formula: "=$A2", want: "=RC1"},
		{name: "range", base: CellRef{Row: 2, Col: 3}, formula: "=SUM(A2:A10)", want: "=SUM(RC[-2]:R[8]C[-2])"},
		{name: "quoted sheet", base: CellRef{Row: 2, Col: 3}, formula: "='売上 集計'!A2", want: "='売上 集計'!RC[-2]"},
		{name: "string only", base: CellRef{Row: 2, Col: 3}, formula: `="A2"`, want: `="A2"`},
		{name: "string comparison", base: CellRef{Row: 2, Col: 3}, formula: `=IF(A2="B2",1,0)`, want: `=IF(RC[-2]="B2",1,0)`},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := NormalizeA1ToR1C1Pattern(tt.formula, NormalizeOptions{BaseCell: tt.base})
			if got.Formula != tt.want {
				t.Fatalf("formula = %q, want %q; result=%#v", got.Formula, tt.want, got)
			}
			if got.Status != ParseStatusOK {
				t.Fatalf("status = %q, want ok; diagnostics=%#v features=%#v", got.Status, got.Diagnostics, got.Features)
			}
		})
	}
}

func TestNormalizeColumnAndRowRanges(t *testing.T) {
	tests := []struct {
		formula string
		want    string
	}{
		{formula: "=SUM(A:A)", want: "=SUM(C[-2]:C[-2])"},
		{formula: "=SUM($A:$C)", want: "=SUM(C1:C3)"},
		{formula: "=SUM(1:1)", want: "=SUM(R[-1]:R[-1])"},
		{formula: "=SUM($1:$10)", want: "=SUM(R1:R10)"},
	}
	for _, tt := range tests {
		got := NormalizeA1ToR1C1Pattern(tt.formula, NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
		if got.Formula != tt.want {
			t.Fatalf("%s => %q, want %q", tt.formula, got.Formula, tt.want)
		}
	}
}

func TestNormalizeSheetQualifiedReferences(t *testing.T) {
	tests := []struct {
		formula string
		want    string
		sheet   string
	}{
		{formula: "=Sheet1!A1", want: "=Sheet1!R[-1]C[-2]", sheet: "Sheet1"},
		{formula: "='O''Brien'!$B$4", want: "='O''Brien'!R4C2", sheet: `'O''Brien'`},
	}
	for _, tt := range tests {
		got := NormalizeA1ToR1C1Pattern(tt.formula, NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
		if got.Formula != tt.want {
			t.Fatalf("%s => %q, want %q", tt.formula, got.Formula, tt.want)
		}
		if len(got.References) != 1 || got.References[0].Sheet != tt.sheet {
			t.Fatalf("references = %#v", got.References)
		}
	}
}

func TestNormalizeDoesNotRewriteFalsePositives(t *testing.T) {
	tests := []string{
		"=LOG10(A1)",
		"=R1C1+A1B2",
		`="A1:B2"`,
		"=#REF!+A1",
	}
	for _, formula := range tests {
		got := NormalizeA1ToR1C1Pattern(formula, NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
		switch formula {
		case "=LOG10(A1)":
			if got.Formula != "=LOG10(R[-1]C[-2])" {
				t.Fatalf("%s => %q", formula, got.Formula)
			}
		case "=#REF!+A1":
			if got.Formula != "=#REF!+R[-1]C[-2]" {
				t.Fatalf("%s => %q", formula, got.Formula)
			}
		default:
			if got.Formula != formula {
				t.Fatalf("%s => %q", formula, got.Formula)
			}
		}
	}
}

func TestNormalizeDetectsP1FeaturesAndPreservesRaw(t *testing.T) {
	tests := []struct {
		name    string
		formula string
		code    string
	}{
		{name: "external", formula: "=[Book.xlsx]Sheet1!A1", code: FeatureExternalReference},
		{name: "quoted external", formula: "='[Book.xlsx]Sheet1'!A1", code: FeatureExternalReference},
		{name: "3d", formula: "=Sheet1:Sheet3!A1", code: Feature3DReference},
		{name: "structured", formula: "=Table1[Amount]*A1", code: FeatureStructuredReference},
		{name: "spill", formula: "=A1#", code: FeatureSpillReference},
		{name: "implicit", formula: "=@A1:A10", code: FeatureImplicitIntersection},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := NormalizeA1ToR1C1Pattern(tt.formula, NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
			if got.Status != ParseStatusPartial {
				t.Fatalf("status = %q, want partial; result=%#v", got.Status, got)
			}
			if !hasFeature(got.Features, tt.code) {
				t.Fatalf("missing feature %q in %#v", tt.code, got.Features)
			}
			switch tt.code {
			case FeatureStructuredReference:
				if got.Formula != "=Table1[Amount]*R[-1]C[-2]" {
					t.Fatalf("formula = %q", got.Formula)
				}
			default:
				if got.Formula != tt.formula {
					t.Fatalf("formula = %q, want preserved %q", got.Formula, tt.formula)
				}
			}
		})
	}
}

func TestNormalizeUnterminatedStructuredReferenceAddsDiagnostic(t *testing.T) {
	got := NormalizeA1ToR1C1Pattern("=Table1[Amount*A1", NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
	if got.Status != ParseStatusPartial {
		t.Fatalf("status = %q, want partial; result=%#v", got.Status, got)
	}
	if got.Formula != "=Table1[Amount*A1" {
		t.Fatalf("formula = %q", got.Formula)
	}
	if !hasFeature(got.Features, FeatureStructuredReference) {
		t.Fatalf("features = %#v", got.Features)
	}
	if !hasDiagnostic(got.Diagnostics, DiagnosticUnterminatedStructuredReference) {
		t.Fatalf("diagnostics = %#v", got.Diagnostics)
	}
}

func TestNormalizeScientificNotationDoesNotCreateReference(t *testing.T) {
	got := NormalizeA1ToR1C1Pattern("=1E10+1.5e-3+A1", NormalizeOptions{BaseCell: CellRef{Row: 2, Col: 3}})
	if got.Formula != "=1E10+1.5e-3+R[-1]C[-2]" {
		t.Fatalf("formula = %q", got.Formula)
	}
	if len(got.References) != 1 || got.References[0].Raw != "A1" {
		t.Fatalf("references = %#v", got.References)
	}
}

func TestNormalizeInvalidBaseCellFails(t *testing.T) {
	got := NormalizeA1ToR1C1Pattern("=A1", NormalizeOptions{BaseCell: CellRef{}})
	if got.Status != ParseStatusFailed {
		t.Fatalf("status = %q, want failed", got.Status)
	}
	if len(got.Diagnostics) != 1 || got.Diagnostics[0].Code != DiagnosticInvalidBaseCell {
		t.Fatalf("diagnostics = %#v", got.Diagnostics)
	}
}

func hasFeature(features []Feature, code string) bool {
	for _, feature := range features {
		if feature.Code == code {
			return true
		}
	}
	return false
}

func hasDiagnostic(diagnostics []Diagnostic, code DiagnosticCode) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}
