package analyze

import (
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

// moduleAnalysisFacts contains the immutable, file-local facts shared by the
// source rules. It is built once for a parsed file revision and is safe to
// retain on parsedFile values copied into worker goroutines.
type moduleAnalysisFacts struct {
	moduleDeclarations  map[string]sourceDeclaration
	procedureLineOwners []int
	procedureDecls      map[int]map[string]sourceDeclaration
}

func buildModuleAnalysisFacts(lines []string, document procedureir.DocumentIR, procedures []sourceProcedure) *moduleAnalysisFacts {
	facts := &moduleAnalysisFacts{
		moduleDeclarations:  make(map[string]sourceDeclaration),
		procedureLineOwners: make([]int, len(lines)+1),
		procedureDecls:      make(map[int]map[string]sourceDeclaration, len(procedures)),
	}
	for line := range facts.procedureLineOwners {
		facts.procedureLineOwners[line] = -1
	}

	ordered := append([]sourceProcedure(nil), procedures...)
	sort.SliceStable(ordered, func(i, j int) bool {
		if ordered[i].StartLine != ordered[j].StartLine {
			return ordered[i].StartLine < ordered[j].StartLine
		}
		return ordered[i].StartByte < ordered[j].StartByte
	})
	coveredThrough := 0
	for index, procedure := range ordered {
		start := procedure.StartLine
		if start < 1 {
			start = 1
		}
		end := procedure.EndLine
		if end > len(lines) {
			end = len(lines)
		}
		if end >= start && end > coveredThrough {
			fillStart := start
			if fillStart <= coveredThrough {
				fillStart = coveredThrough + 1
			}
			for line := fillStart; line <= end; line++ {
				facts.procedureLineOwners[line] = index
			}
			coveredThrough = end
		}
		facts.procedureDecls[procedure.StartByte] = procedureDeclarations(lines, procedure)
	}

	if !document.Parse.HasError && !document.Parse.HasMissing && len(document.Declarations) > 0 {
		for _, declaration := range document.Declarations {
			if !isSourceModuleDeclaration(declaration) {
				continue
			}
			if converted, ok := sourceDeclarationFromIR(declaration); ok {
				facts.moduleDeclarations[strings.ToLower(converted.Name)] = converted
			}
		}
	} else {
		facts.moduleDeclarations = moduleDeclarationsFromSource(lines, facts.procedureLineOwners)
	}
	return facts
}

func isSourceModuleDeclaration(declaration procedureir.Declaration) bool {
	if declaration.Scope != procedureir.ScopeModule && declaration.Scope != procedureir.ScopeProject {
		return false
	}
	switch strings.ToLower(strings.TrimSpace(declaration.Kind)) {
	case "variable", "module_variable", "field", "withevents_field", "variable_declaration", "const", "constant", "const_declaration":
		return true
	default:
		return false
	}
}

func sourceDeclarationFromIR(declaration procedureir.Declaration) (sourceDeclaration, bool) {
	name := strings.TrimSpace(declaration.Name)
	if name == "" {
		return sourceDeclaration{}, false
	}
	return sourceDeclaration{
		Name:          name,
		Type:          declaration.Type,
		Line:          declaration.Range.StartLine,
		Object:        declaration.IsObject,
		Array:         declaration.IsArray,
		Fixed:         declaration.ValueShape == procedureir.ValueShapeFixedArray,
		NewExpression: strings.HasPrefix(strings.ToLower(strings.TrimSpace(declaration.Type)), "new "),
	}, true
}

func (facts *moduleAnalysisFacts) lineInProcedure(line int) bool {
	return facts != nil && line >= 0 && line < len(facts.procedureLineOwners) && facts.procedureLineOwners[line] >= 0
}

func (facts *moduleAnalysisFacts) procedureDeclarations(procedure sourceProcedure) map[string]sourceDeclaration {
	if facts == nil {
		return nil
	}
	return facts.procedureDecls[procedure.StartByte]
}

func moduleDeclarationsFromSource(lines []string, procedureLineOwners []int) map[string]sourceDeclaration {
	decls := make(map[string]sourceDeclaration)
	for lineNo, rawLine := range lines {
		line := lineNo + 1
		if line >= 0 && line < len(procedureLineOwners) && procedureLineOwners[line] >= 0 {
			continue
		}
		stmt := normalizedCodeLine(rawLine)
		lower := strings.ToLower(stmt)
		if !strings.HasPrefix(lower, "dim ") && !strings.HasPrefix(lower, "static ") && !strings.HasPrefix(lower, "private ") && !strings.HasPrefix(lower, "public ") {
			continue
		}
		match := declRe.FindStringSubmatch(stmt)
		if len(match) == 0 {
			continue
		}
		for _, part := range splitArgs(match[1]) {
			name, typ, array, newExpression := declarationNameAndType(part)
			if name == "" {
				continue
			}
			decls[strings.ToLower(name)] = sourceDeclaration{
				Name: name, Type: typ, Line: line, Object: isObjectType(typ),
				Array: array, NewExpression: newExpression,
			}
		}
	}
	return decls
}

func (file parsedFile) moduleAnalysisFacts() *moduleAnalysisFacts {
	if file.ModuleFacts != nil {
		return file.ModuleFacts
	}
	return buildModuleAnalysisFacts(file.Lines, file.IR, file.procedureProjection())
}

func (file parsedFile) procedureDeclarationsFor(proc sourceProcedure) map[string]sourceDeclaration {
	if facts := file.moduleAnalysisFacts(); facts != nil {
		if declarations := facts.procedureDeclarations(proc); declarations != nil {
			return declarations
		}
	}
	return procedureDeclarations(file.Lines, proc)
}
