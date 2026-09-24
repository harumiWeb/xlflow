package analyze

import (
	"cmp"
	"slices"
	"strconv"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

const (
	excelMaxRows    = 1048576
	excelMaxColumns = 16384 // XFD
)

// udfCellReferenceFindings reports worksheet-visible Functions in standard
// modules whose names are valid Excel cell references. A public Function in a
// standard module is exposed to Excel as a worksheet UDF; when the function
// name parses as a cell reference, worksheet formulas resolve the cell
// instead of calling the function. The scan is declaration-level only: it
// covers A1-style references within the current worksheet limits and absolute
// R1C1-style references, and skips Private/Friend members, non-Function
// procedures, Option Private Module files, and non-standard module kinds that
// can never produce a worksheet UDF.
func (a Analyzer) udfCellReferenceFindings(file parsedFile) []Finding {
	if file.IR.Parse.HasError || file.IR.Parse.HasMissing {
		return nil
	}
	if !strings.EqualFold(file.ModuleKind, "standard") {
		return nil
	}
	if file.moduleAnalysisFacts().privateModulePresent() {
		return nil
	}
	var findings []Finding
	for _, procedure := range file.IR.Procedures {
		symbol := procedure.Symbol
		if symbol.Kind != procedureir.ProcedureFunction || symbol.Recovered || len(symbol.ConditionalBranches) > 0 {
			continue
		}
		if visibility := strings.TrimSpace(symbol.Visibility); visibility != "" && !strings.EqualFold(visibility, "public") {
			continue
		}
		name := cleanIdentifier(symbol.Name)
		style, ok := udfCellReferenceStyle(name)
		if !ok {
			continue
		}
		finding := a.simpleFinding(file, sourceProcedure{Name: symbol.Name}, symbol.DeclarationRange.StartLine, "VBA270", "warning",
			"Public Function "+name+" is exposed to Excel worksheets as a UDF, but its name is a valid "+style+" cell reference.",
			"Worksheet formulas resolve the cell reference instead of calling the function, so the UDF cannot be invoked from a cell.",
			"Rename the function so it cannot be parsed as a cell reference, or hide it from worksheets by declaring it Private or adding Option Private Module.")
		finding.Column = symbol.DeclarationRange.StartColumn
		finding.EndLine = symbol.DeclarationRange.EndLine
		finding.EndColumn = symbol.DeclarationRange.EndColumn
		findings = append(findings, finding)
	}
	slices.SortFunc(findings, func(x, y Finding) int {
		return cmp.Or(cmp.Compare(x.Line, y.Line), cmp.Compare(x.Column, y.Column))
	})
	return findings
}

// udfCellReferenceStyle reports whether name is a valid Excel cell reference
// in A1 or R1C1 notation, returning the matched style label for diagnostics.
// A1 references use the current worksheet limits (columns through XFD, rows
// through 1048576). R1C1 coverage is restricted to absolute R<row>C<column>
// cell references; bare RC and row-only/column-only R1C1 forms such as R5 or
// C3 are range references rather than cell references, though names of that
// shape still collide as ordinary A1 references and are reported there.
func udfCellReferenceStyle(name string) (string, bool) {
	if a1CellReference(name) {
		return "A1-style", true
	}
	if r1c1CellReference(name) {
		return "R1C1-style", true
	}
	return "", false
}

func a1CellReference(name string) bool {
	i := 0
	for i < len(name) && isASCIILetter(name[i]) {
		i++
	}
	if i == 0 || i > 3 || i == len(name) {
		return false
	}
	if _, ok := excelColumnIndex(name[:i]); !ok {
		return false
	}
	row, ok := excelReferenceNumber(name[i:])
	return ok && row >= 1 && row <= excelMaxRows
}

func r1c1CellReference(name string) bool {
	if len(name) < 4 || (name[0] != 'R' && name[0] != 'r') {
		return false
	}
	i := 1
	for i < len(name) && isASCIIDigit(name[i]) {
		i++
	}
	if i == 1 || i >= len(name) || (name[i] != 'C' && name[i] != 'c') {
		return false
	}
	row, rowOK := excelReferenceNumber(name[1:i])
	column, columnOK := excelReferenceNumber(name[i+1:])
	return rowOK && columnOK &&
		row >= 1 && row <= excelMaxRows &&
		column >= 1 && column <= excelMaxColumns
}

// excelColumnIndex converts 1-3 ASCII column letters to a 1-based column
// index and validates it against the worksheet column limit.
func excelColumnIndex(letters string) (int, bool) {
	if len(letters) == 0 || len(letters) > 3 {
		return 0, false
	}
	index := 0
	for i := 0; i < len(letters); i++ {
		c := letters[i]
		if c >= 'a' && c <= 'z' {
			c -= 'a' - 'A'
		}
		if c < 'A' || c > 'Z' {
			return 0, false
		}
		index = index*26 + int(c-'A') + 1
	}
	return index, index <= excelMaxColumns
}

// excelReferenceNumber parses a nonempty ASCII digit run into a positive
// integer. Leading zeros are accepted because Excel resolves forms such as
// =A01 to the same cell.
func excelReferenceNumber(digits string) (int, bool) {
	if digits == "" {
		return 0, false
	}
	for i := 0; i < len(digits); i++ {
		if !isASCIIDigit(digits[i]) {
			return 0, false
		}
	}
	value, err := strconv.Atoi(digits)
	if err != nil {
		return 0, false
	}
	return value, true
}

func isASCIILetter(c byte) bool {
	return c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z'
}

func isASCIIDigit(c byte) bool {
	return c >= '0' && c <= '9'
}
