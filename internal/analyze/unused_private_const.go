package analyze

import (
	"sort"
	"strings"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// unusedPrivateConstFindings reports module-level Private Const declarations
// that are never referenced. A Private Const is file-local by language rules,
// so a single-file scan is complete: references can only come from procedure
// bodies, module-level declaration initializers and attributes, conditional
// compilation directives, or a raw textual mention (comments and string
// literals mark the name used, which only suppresses findings).
func (a Analyzer) unusedPrivateConstFindings(file parsedFile) []Finding {
	if file.IR.Parse.HasError || file.IR.Parse.HasMissing {
		return nil
	}
	var consts []procedureir.Declaration
	for _, declaration := range file.IR.Declarations {
		// Enum members carry IsConst because they are compile-time constants,
		// but they are not Const declarations and are referenced through the
		// enum's qualified name; they are out of scope for this rule.
		if !declaration.IsConst || declaration.Kind != "const" || declaration.Parent != "" ||
			!strings.EqualFold(strings.TrimSpace(declaration.Visibility), "private") ||
			declaration.Recovered || len(declaration.ConditionalBranches) > 0 {
			continue
		}
		if name := cleanIdentifier(declaration.Name); name == "" || isIgnoredUnusedName(name) {
			continue
		}
		consts = append(consts, declaration)
	}
	if len(consts) == 0 {
		return nil
	}

	// shadowedByProcedure indexes procedures that declare a local or parameter
	// with the same name; occurrences inside those bodies bind to the local.
	shadowed := make(map[string]map[int]bool)
	procedureRanges := make([]vbaast.Range, 0, len(file.IR.Procedures))
	for _, procedure := range file.IR.Procedures {
		bodyRange := procedure.Symbol.BodyRange
		if bodyRange.EndByte <= bodyRange.StartByte {
			bodyRange = procedure.Symbol.DeclarationRange
		}
		procedureRanges = append(procedureRanges, bodyRange)
		for _, declaration := range procedure.Declarations {
			if declaration.Scope != procedureir.ScopeLocal && declaration.Scope != procedureir.ScopeParameter {
				continue
			}
			name := assignmentCanonicalName(declaration.Name)
			if shadowed[name] == nil {
				shadowed[name] = make(map[int]bool)
			}
			shadowed[name][bodyRange.StartByte] = true
		}
	}

	var findings []Finding
	for _, declaration := range consts {
		name := assignmentCanonicalName(declaration.Name)
		if privateConstReferenced(file, name, declaration.Range, procedureRanges, shadowed[name]) {
			continue
		}
		finding := a.simpleFinding(file, sourceProcedure{}, declaration.Range.StartLine, "VBA261", "warning",
			"Private Const "+cleanIdentifier(declaration.Name)+" is never referenced.",
			"No procedure access, declaration initializer, or conditional-compilation directive in this module uses the constant.",
			"Remove the declaration, or reference it where the value is needed.")
		finding.Column = declaration.Range.StartColumn + 1
		finding.EndLine = declaration.Range.EndLine
		finding.EndColumn = declaration.Range.EndColumn + 1
		findings = append(findings, finding)
	}
	sort.Slice(findings, func(i, j int) bool {
		if findings[i].Line != findings[j].Line {
			return findings[i].Line < findings[j].Line
		}
		return findings[i].Column < findings[j].Column
	})
	return findings
}

// privateConstReferenced reports whether a whole-word occurrence of the
// constant name exists outside its own declarator and outside procedure
// bodies where a local declaration shadows it.
func privateConstReferenced(file parsedFile, name string, declarator vbaast.Range, procedureRanges []vbaast.Range, shadowed map[int]bool) bool {
	matcher := unusedNameMatcher(name)
	for i, line := range file.Lines {
		lineNo := i + 1
		for _, loc := range matcher.FindAllStringIndex(line, -1) {
			begin := loc[0]
			if !isIdentifierByte(line[begin]) {
				begin++
			}
			if positionInRange(lineNo, begin+1, declarator) {
				continue
			}
			shadowedHere := false
			for _, procRange := range procedureRanges {
				if !positionInRange(lineNo, begin+1, procRange) {
					continue
				}
				if shadowed[procRange.StartByte] {
					shadowedHere = true
				}
				break
			}
			if shadowedHere {
				continue
			}
			return true
		}
	}
	return false
}
