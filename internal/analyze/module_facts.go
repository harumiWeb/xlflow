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
	// privateModule records the module option once for this source revision.
	// The zero value is unknown so incomplete/recovered syntax fails open.
	privateModule moduleOptionState

	// arrayOperationsByName and arrayOperationsByLine are immutable indexes over
	// the small set of module-wide operations used by array interprocedural
	// summaries. Values are only exposed through callback accessors below; in
	// particular, callers cannot retain or mutate the backing slices.
	arrayOperationsByName map[string][]moduleArrayOperationFact
	arrayOperationsByLine map[int][]moduleArrayOperationFact
	lineFacts             []moduleSourceLineFact
}

type moduleOptionState uint8

const (
	moduleOptionUnknown moduleOptionState = iota
	moduleOptionAbsent
	moduleOptionPresent
)

type moduleArrayOperationKind uint8

const (
	moduleArrayWholeAssignment moduleArrayOperationKind = iota + 1
	moduleArrayDirectRedim
	moduleArrayErase
)

type moduleArrayOperationFact struct {
	Name       string
	Line       int
	RHS        string
	Dimensions string
	Preserve   bool
	Kind       moduleArrayOperationKind
}

type moduleSourceLineFact struct {
	executable     bool
	setupGuardName string
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
	facts.privateModule = buildModuleSourceFacts(lines, facts, document.Parse.HasError || document.Parse.HasMissing)
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

// buildModuleSourceFacts performs the one source-line pass needed by the
// module-wide array summaries. normalizedCodeLine is intentionally called only
// here for these facts; procedure consumers read the resulting immutable
// operation/index values through accessors instead of repeating the scan.
func buildModuleSourceFacts(lines []string, facts *moduleAnalysisFacts, parseIncomplete bool) moduleOptionState {
	if facts == nil {
		return moduleOptionUnknown
	}
	privateModule := moduleOptionAbsent
	privateModuleCandidate := false
	for lineNo, rawLine := range lines {
		candidate := moduleFactCandidateLine(rawLine)
		// Before a setup guard is found, only lines that can contribute to the
		// indexed facts need normalization. Once lineFacts is materialized, the
		// remainder of the procedure must also be normalized so the final
		// executable-line check preserves the legacy all-line behavior.
		if !candidate && facts.lineFacts == nil {
			continue
		}
		text := normalizedCodeLine(rawLine)
		lower := strings.ToLower(strings.TrimSpace(text))
		if strings.HasPrefix(lower, "if ") {
			match := arraySetupGuardRe.FindStringSubmatch(text)
			if len(match) == 2 {
				facts.ensureLineFacts(lines, lineNo)
				facts.lineFacts[lineNo].setupGuardName = strings.ToLower(cleanIdentifier(match[1]))
			}
		}
		if facts.lineFacts != nil {
			facts.lineFacts[lineNo].executable = lower != "" && lower != "end sub" && lower != "end function" && lower != "end property"
		}
		if !candidate {
			continue
		}
		if strings.Contains(lower, "=") {
			if lhs, rhs, indexed, assigned := arrayAssignment(text); assigned && !indexed {
				facts.addModuleArrayOperation(moduleArrayOperationFact{
					Name: strings.ToLower(cleanIdentifier(lhs)), Line: lineNo,
					RHS: strings.TrimSpace(rhs), Kind: moduleArrayWholeAssignment,
				})
			}
		}
		if strings.HasPrefix(lower, "redim ") {
			match := arrayRedimRe.FindStringSubmatch(text)
			if len(match) > 0 {
				preserve := strings.TrimSpace(match[1]) != ""
				for _, clause := range splitArgs(match[2]) {
					redim, direct := parseDirectArrayRedimClause(clause)
					if !direct {
						continue
					}
					facts.addModuleArrayOperation(moduleArrayOperationFact{
						Name: strings.ToLower(cleanIdentifier(redim.name)), Line: lineNo,
						Dimensions: redim.dimensions, Preserve: preserve,
						Kind: moduleArrayDirectRedim,
					})
				}
			}
		}
		if strings.HasPrefix(lower, "erase ") {
			match := arrayEraseRe.FindStringSubmatch(text)
			if len(match) == 2 {
				for _, target := range splitArgs(match[1]) {
					name := strings.ToLower(strings.TrimSpace(target))
					if !arrayEraseNameRe.MatchString(name) {
						continue
					}
					facts.addModuleArrayOperation(moduleArrayOperationFact{
						Name: name, Line: lineNo, Kind: moduleArrayErase,
					})
				}
			}
		}
		switch lower {
		case "option private module":
			privateModule = moduleOptionPresent
		case "option private":
			privateModuleCandidate = true
		default:
			if strings.HasPrefix(lower, "option private ") {
				privateModuleCandidate = true
			}
		}
	}
	if privateModuleCandidate || parseIncomplete {
		return moduleOptionUnknown
	}
	return privateModule
}

// moduleFactCandidateLine is a cheap lexical gate for the small subset of
// source lines that can contribute module options or array operation facts.
// Most lines in ordinary modules are declarations or procedure boundaries;
// avoiding normalization for those lines keeps the shared-facts fast path
// allocation-light without changing the conservative parser-backed matching
// below. False positives are harmless because the full normalizer still runs.
func moduleFactCandidateLine(line string) bool {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" {
		return false
	}
	if asciiFoldPrefix(trimmed, "option") ||
		asciiFoldKeywordPrefix(trimmed, "if") ||
		asciiFoldKeywordPrefix(trimmed, "redim") ||
		asciiFoldKeywordPrefix(trimmed, "erase") {
		return true
	}
	return strings.Contains(trimmed, "=")
}

func asciiFoldPrefix(value, prefix string) bool {
	if len(value) < len(prefix) {
		return false
	}
	for index := range prefix {
		left, right := value[index], prefix[index]
		if left >= 'A' && left <= 'Z' {
			left += 'a' - 'A'
		}
		if right >= 'A' && right <= 'Z' {
			right += 'a' - 'A'
		}
		if left != right {
			return false
		}
	}
	return true
}

func asciiFoldKeywordPrefix(value, keyword string) bool {
	return asciiFoldPrefix(value, keyword) && len(value) > len(keyword) &&
		(value[len(keyword)] == ' ' || value[len(keyword)] == '\t')
}

func (facts *moduleAnalysisFacts) addModuleArrayOperation(operation moduleArrayOperationFact) {
	if facts == nil || operation.Name == "" {
		return
	}
	if facts.arrayOperationsByName == nil {
		facts.arrayOperationsByName = make(map[string][]moduleArrayOperationFact)
		facts.arrayOperationsByLine = make(map[int][]moduleArrayOperationFact)
	}
	facts.arrayOperationsByName[operation.Name] = append(facts.arrayOperationsByName[operation.Name], operation)
	facts.arrayOperationsByLine[operation.Line] = append(facts.arrayOperationsByLine[operation.Line], operation)
}

func (facts *moduleAnalysisFacts) ensureLineFacts(lines []string, through int) {
	if facts == nil || facts.lineFacts != nil {
		return
	}
	facts.lineFacts = make([]moduleSourceLineFact, len(lines))
	if through >= len(lines) {
		through = len(lines) - 1
	}
	for lineNo := 0; lineNo <= through; lineNo++ {
		lower := strings.ToLower(strings.TrimSpace(normalizedCodeLine(lines[lineNo])))
		facts.lineFacts[lineNo].executable = lower != "" && lower != "end sub" && lower != "end function" && lower != "end property"
	}
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

// forEachConstantForProcedure visits the constants visible from one
// procedure. Module constants are visited first, followed by local constants
// in source order, so a procedure-local declaration correctly shadows a
// same-named module constant for consumers that build an environment map.
func (facts *moduleAnalysisFacts) forEachConstantForProcedure(procedure sourceProcedure, visit func(moduleConstantFact)) {
	if facts == nil || visit == nil {
		return
	}
	for _, constant := range facts.constants {
		if constant.Module {
			visit(constant)
		}
	}
	for _, constant := range facts.constants {
		if !constant.Module && constant.Line >= procedure.StartLine && constant.Line <= procedure.EndLine {
			visit(constant)
		}
	}
}

// forEachSourceConstantForProcedure is the no-facts compatibility path for
// standalone callers. It retains module constants and current-procedure local
// constants while excluding locals owned by another procedure.
func forEachSourceConstantForProcedure(file parsedFile, procedure sourceProcedure, visit func(moduleConstantFact)) {
	if visit == nil {
		return
	}
	type procedureRange struct{ start, end int }
	ranges := make([]procedureRange, 0, len(file.IR.Procedures))
	for _, candidate := range file.IR.Procedures {
		start := candidate.Symbol.DeclarationRange.StartLine
		end := candidate.Symbol.DeclarationRange.EndLine
		if start > 0 && end >= start {
			ranges = append(ranges, procedureRange{start: start, end: end})
		}
	}
	var moduleConstants []moduleConstantFact
	var localConstants []moduleConstantFact
	for lineIndex, rawLine := range file.Lines {
		name, expression, ok := fileConstDeclaration(rawLine)
		if !ok || strings.TrimSpace(name) == "" {
			continue
		}
		line := lineIndex + 1
		ownedByOther := false
		ownedByCurrent := false
		for _, candidate := range ranges {
			if line < candidate.start || line > candidate.end {
				continue
			}
			if candidate.start == procedure.StartLine && candidate.end == procedure.EndLine {
				ownedByCurrent = true
			} else {
				ownedByOther = true
			}
			break
		}
		if ownedByOther {
			continue
		}
		constant := moduleConstantFact{
			Name: strings.TrimSpace(name), Expression: strings.TrimSpace(expression), Line: line,
			Module: !ownedByCurrent,
		}
		if ownedByCurrent || (len(ranges) == 0 && line >= procedure.StartLine && line <= procedure.EndLine) {
			localConstants = append(localConstants, constant)
		} else {
			moduleConstants = append(moduleConstants, constant)
		}
	}
	for _, constant := range moduleConstants {
		visit(constant)
	}
	for _, constant := range localConstants {
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

func (facts *moduleAnalysisFacts) privateModuleState() moduleOptionState {
	if facts == nil {
		return moduleOptionUnknown
	}
	return facts.privateModule
}

func (facts *moduleAnalysisFacts) privateModulePresent() bool {
	return facts.privateModuleState() == moduleOptionPresent
}

// forEachArrayOperationFor visits a copy of each operation for one canonical
// identifier. The callback API deliberately keeps the immutable backing slice
// private to moduleAnalysisFacts.
func (facts *moduleAnalysisFacts) forEachArrayOperationFor(name string, visit func(moduleArrayOperationFact)) {
	if facts == nil || visit == nil {
		return
	}
	for _, operation := range facts.arrayOperationsByName[strings.ToLower(cleanIdentifier(name))] {
		visit(operation)
	}
}

func (facts *moduleAnalysisFacts) forEachArrayOperationAt(line int, visit func(moduleArrayOperationFact)) {
	if facts == nil || visit == nil {
		return
	}
	for _, operation := range facts.arrayOperationsByLine[line] {
		visit(operation)
	}
}

func (facts *moduleAnalysisFacts) sourceLineIsExecutable(line int) bool {
	return facts != nil && line >= 0 && line < len(facts.lineFacts) && facts.lineFacts[line].executable
}

func (facts *moduleAnalysisFacts) sourceLineSetupGuard(line int) (string, bool) {
	if facts == nil || line < 0 || line >= len(facts.lineFacts) {
		return "", false
	}
	name := facts.lineFacts[line].setupGuardName
	return name, name != ""
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

// ensureModuleAnalysisFacts attaches the immutable facts object to the
// caller-owned parsed-file revision. Normal batch/realtime setup attaches it
// before worker launch; this helper is the explicit compatibility boundary for
// package-local synthetic callers.
func (file *parsedFile) ensureModuleAnalysisFacts() *moduleAnalysisFacts {
	if file == nil {
		return nil
	}
	if file.ModuleFacts == nil {
		file.ModuleFacts = buildModuleAnalysisFacts(file.Lines, file.IR, file.procedureProjection())
		if file.ModuleDeclarations == nil && file.ModuleFacts != nil {
			file.ModuleDeclarations = file.ModuleFacts.moduleDeclarations
		}
	}
	return file.ModuleFacts
}

func (file *parsedFile) moduleAnalysisFacts() *moduleAnalysisFacts {
	return file.ensureModuleAnalysisFacts()
}

func (file parsedFile) procedureDeclarationsFor(proc sourceProcedure) map[string]sourceDeclaration {
	if facts := file.moduleAnalysisFacts(); facts != nil {
		if declarations := facts.procedureDeclarations(proc); declarations != nil {
			return declarations
		}
	}
	return procedureDeclarations(file.Lines, proc)
}
