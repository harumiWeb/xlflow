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
	// procedureFactsByStart is a compact declaration-start index to the
	// already-owned procedure facts. It stores pointers, not another copy of
	// procedure IR, so module consumers can resolve a procedure revision
	// without rebuilding its indexes.
	procedureFactsByStart map[int]*procedureAnalysisFacts
	// constants preserves source order because constant expressions may refer
	// to declarations that appeared earlier in the module. The backing slice
	// is owned by the facts object and is only exposed through read-only
	// iteration helpers below.
	constants              []moduleConstantFact
	constantNames          map[string]struct{}
	moduleConstantNames    map[string]struct{}
	procedureConstantNames map[int]map[string]struct{}
	procedureNames         map[string]struct{}
}

type moduleConstantFact struct {
	Name       string
	Expression string
	Line       int
	// Module is false for a procedure-local Const. Keeping this bit alongside
	// the ordered declaration avoids re-scanning source/IR for consumers that
	// need to distinguish module constants from local constants.
	Module bool
}

func buildModuleAnalysisFacts(lines []string, document procedureir.DocumentIR, procedures []sourceProcedure) *moduleAnalysisFacts {
	facts := &moduleAnalysisFacts{
		moduleDeclarations:  make(map[string]sourceDeclaration),
		procedureLineOwners: make([]int, len(lines)+1),
		procedureDecls:      make(map[int]map[string]sourceDeclaration, len(procedures)),
	}
	if len(procedures) > 0 {
		facts.procedureFactsByStart = make(map[int]*procedureAnalysisFacts, len(procedures))
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
		procedureFacts := procedure.Facts
		if procedureFacts == nil {
			// Compatibility callers may provide a hand-built source projection
			// without the normal sourceProceduresFromIR attachment. Build once for
			// this module index rather than once per lookup.
			procedureFacts = procedure.analysisFacts()
		}
		facts.procedureFactsByStart[procedure.StartByte] = procedureFacts
		if name := strings.ToLower(strings.TrimSpace(procedure.Name)); name != "" {
			if facts.procedureNames == nil {
				facts.procedureNames = make(map[string]struct{}, len(procedures))
			}
			facts.procedureNames[name] = struct{}{}
		}
	}
	// Some package-local callers build facts from a DocumentIR while passing a
	// minimal sourceProcedure projection. Retain the IR names as a fallback so
	// the same-module index remains complete for those revisions too.
	for _, procedure := range document.Procedures {
		if name := strings.ToLower(strings.TrimSpace(procedure.Symbol.Name)); name != "" {
			if facts.procedureNames == nil {
				facts.procedureNames = make(map[string]struct{}, len(document.Procedures))
			}
			facts.procedureNames[name] = struct{}{}
		}
	}
	for lineNo, rawLine := range lines {
		name, expression, ok := fileConstDeclaration(rawLine)
		if !ok {
			continue
		}
		name = strings.TrimSpace(name)
		if name == "" {
			continue
		}
		facts.constants = append(facts.constants, moduleConstantFact{
			Name:       name,
			Expression: strings.TrimSpace(expression),
			Line:       lineNo + 1,
			Module:     lineNo+1 >= 0 && lineNo+1 < len(facts.procedureLineOwners) && facts.procedureLineOwners[lineNo+1] < 0,
		})
		if facts.constantNames == nil {
			facts.constantNames = make(map[string]struct{})
		}
		key := strings.ToLower(name)
		facts.constantNames[key] = struct{}{}
		owner := facts.procedureLineOwners[lineNo+1]
		if owner < 0 || owner >= len(ordered) {
			if facts.moduleConstantNames == nil {
				facts.moduleConstantNames = make(map[string]struct{})
			}
			facts.moduleConstantNames[key] = struct{}{}
			continue
		}
		if facts.procedureConstantNames == nil {
			facts.procedureConstantNames = make(map[int]map[string]struct{})
		}
		localNames := facts.procedureConstantNames[ordered[owner].StartByte]
		if localNames == nil {
			localNames = make(map[string]struct{})
			facts.procedureConstantNames[ordered[owner].StartByte] = localNames
		}
		localNames[key] = struct{}{}
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
	typ := strings.TrimSpace(declaration.Type)
	newExpression := strings.HasPrefix(strings.ToLower(typ), "new ")
	if newExpression {
		typ = strings.TrimSpace(typ[4:])
	}
	return sourceDeclaration{
		Name: name,
		Type: typ,
		Line: declaration.Range.StartLine,
		// Preserve known object types and IR object classifications, except for
		// MSForms controls. Form controls are initialized by the form runtime and
		// are not reliable roots for this rule; treating them as ordinary nullable
		// objects caused false positives in the corpus.
		Object:        sourceDeclarationIsObject(declaration, typ),
		Array:         declaration.IsArray,
		Fixed:         declaration.ValueShape == procedureir.ValueShapeFixedArray,
		NewExpression: newExpression,
	}, true
}

func sourceDeclarationIsObject(declaration procedureir.Declaration, typ string) bool {
	return (declaration.IsObject || isObjectType(typ)) && !isMSFormsControlType(typ)
}

func isMSFormsControlType(typ string) bool {
	typ = strings.ToLower(strings.TrimSpace(typ))
	return strings.HasPrefix(typ, "msforms.")
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

// procedureFactsFor resolves the immutable facts associated with a procedure
// declaration start byte. The sourceProcedure fallback keeps package-local
// synthetic callers compatible when they bypass module setup.
func (facts *moduleAnalysisFacts) procedureFactsFor(procedure sourceProcedure) *procedureAnalysisFacts {
	if facts != nil {
		if procedureFacts := facts.procedureFactsByStart[procedure.StartByte]; procedureFacts != nil {
			return procedureFacts
		}
	}
	return procedure.analysisFacts()
}

// forEachConstant visits declarations in source order. Callers must treat the
// value as immutable; the callback is deliberately the only public-in-package
// access path to the facts slice.
func (facts *moduleAnalysisFacts) forEachConstant(visit func(moduleConstantFact)) {
	if facts == nil || visit == nil {
		return
	}
	for _, constant := range facts.constants {
		visit(constant)
	}
}

func (facts *moduleAnalysisFacts) hasConstant(name string) bool {
	if facts == nil {
		return false
	}
	_, ok := facts.constantNames[strings.ToLower(strings.TrimSpace(name))]
	return ok
}

// hasConstantForProcedure reports constants visible from one procedure: all
// module constants plus only the Const declarations owned by that procedure.
// Procedure-local names are indexed by declaration start byte so repeated
// rule-family lookups do not rescan the module's ordered constant slice.
func (facts *moduleAnalysisFacts) hasConstantForProcedure(name string, procedure sourceProcedure) bool {
	if facts == nil {
		return false
	}
	key := strings.ToLower(strings.TrimSpace(name))
	if _, ok := facts.moduleConstantNames[key]; ok {
		return true
	}
	localNames := facts.procedureConstantNames[procedure.StartByte]
	_, ok := localNames[key]
	return ok
}

func (facts *moduleAnalysisFacts) hasProcedure(name string) bool {
	if facts == nil {
		return false
	}
	_, ok := facts.procedureNames[strings.ToLower(strings.TrimSpace(name))]
	return ok
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
