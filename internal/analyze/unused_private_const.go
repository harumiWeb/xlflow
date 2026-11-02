package analyze

import (
	"sort"
	"strings"
	"unicode"
	"unicode/utf8"

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
	if file.AuthoredSource != nil {
		file.Lines = normalizedSourceLines(string(file.AuthoredSource))
	}
	constantNames := make([]string, 0, len(consts))
	for _, declaration := range consts {
		constantNames = append(constantNames, assignmentCanonicalName(declaration.Name))
	}
	lexicalIndex := buildUnusedPrivateConstLexicalIndex(file.Lines, constantNames)

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

	moduleName := strings.TrimSpace(file.IR.ModuleName)
	if moduleName == "" {
		moduleName = file.Module
	}
	var findings []Finding
	for _, declaration := range consts {
		name := assignmentCanonicalName(declaration.Name)
		if privateConstReferencedIndexed(file.Lines, lexicalIndex, moduleName, name, declaration.Range, procedureRanges, shadowed[name]) {
			continue
		}
		finding := a.simpleFinding(file, sourceProcedure{}, declaration.Range.StartLine, "VBA266", "warning",
			"Private Const "+cleanIdentifier(declaration.Name)+" is never referenced.",
			"No procedure access, declaration initializer, or conditional-compilation directive in this module uses the constant.",
			"Remove the declaration, or reference it where the value is needed.")
		finding.Column = declaration.Range.StartColumn
		finding.EndLine = declaration.Range.EndLine
		finding.EndColumn = declaration.Range.EndColumn
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
// bodies where a local declaration shadows it. A local shadows only
// unqualified references: `Module.Const` still binds to the module constant,
// so a qualified occurrence inside a shadowing procedure remains a use.
func privateConstReferencedIndexed(lines []string, index *unusedPrivateConstLexicalIndex, moduleName, name string, declarator vbaast.Range, procedureRanges []vbaast.Range, shadowed map[int]bool) bool {
	if index == nil {
		return false
	}
	for _, occurrence := range index.byName[name] {
		line := lines[occurrence.line-1]
		if positionInRange(occurrence.line, occurrence.begin+1, declarator) {
			continue
		}
		if len(shadowed) == 0 {
			return true
		}
		shadowedHere := false
		for _, procRange := range procedureRanges {
			if !positionInRange(occurrence.line, occurrence.begin+1, procRange) {
				continue
			}
			if shadowed[procRange.StartByte] {
				shadowedHere = true
			}
			break
		}
		if shadowedHere && !qualifiedModuleReference(line[:occurrence.begin], moduleName) {
			continue
		}
		return true
	}
	return false
}

type unusedPrivateConstOccurrence struct {
	line      int
	begin     int
	matchFrom int
	matchTo   int
}

type unusedPrivateConstLexicalIndexStats struct {
	sourceLinesScanned    int
	asciiTokensScanned    int
	fallbackRegexLineScan int
}

// unusedPrivateConstLexicalIndex stores authored-source occurrences by
// candidate name. Common ASCII identifiers are found by one source pass;
// unusual names and rare Unicode case-fold equivalents retain regexp behavior.
type unusedPrivateConstLexicalIndex struct {
	byName map[string][]unusedPrivateConstOccurrence
	stats  unusedPrivateConstLexicalIndexStats
}

func buildUnusedPrivateConstLexicalIndex(lines, names []string) *unusedPrivateConstLexicalIndex {
	index := &unusedPrivateConstLexicalIndex{byName: make(map[string][]unusedPrivateConstOccurrence, len(names))}
	fastNames := make(map[uint64][]string, len(names))
	seenNames := make(map[string]struct{}, len(names))
	candidateNames := make([]string, 0, len(names))
	for _, name := range names {
		if name == "" {
			continue
		}
		if _, exists := seenNames[name]; exists {
			continue
		}
		seenNames[name] = struct{}{}
		candidateNames = append(candidateNames, name)
		index.byName[name] = nil
		if !unusedPrivateConstASCIIIdentifier(name) {
			continue
		}
		key := unusedPrivateConstASCIIIdentifierHash(name)
		fastNames[key] = append(fastNames[key], name)
	}

	hasUnicodeASCIIFold := false
	for lineNumber, line := range lines {
		index.stats.sourceLinesScanned++
		for position := 0; position < len(line); {
			if unusedPrivateConstASCIIIdentifierByte(line[position]) {
				start := position
				for position < len(line) && unusedPrivateConstASCIIIdentifierByte(line[position]) {
					position++
				}
				index.stats.asciiTokensScanned++
				token := line[start:position]
				for _, name := range fastNames[unusedPrivateConstASCIIIdentifierHash(token)] {
					if !unusedPrivateConstASCIIEqualFold(token, name) {
						continue
					}
					index.byName[name] = append(index.byName[name], unusedPrivateConstTokenOccurrence(line, lineNumber+1, start, position))
				}
				continue
			}
			if line[position] >= utf8.RuneSelf {
				r, size := utf8.DecodeRuneInString(line[position:])
				if _, ok := unusedPrivateConstUnicodeFoldASCII[r]; ok {
					hasUnicodeASCIIFold = true
				}
				position += size
				continue
			}
			position++
		}
	}

	for _, name := range candidateNames {
		// The regexp's case-insensitive flag also applies to its negated
		// ASCII boundary classes: Kelvin sign and long s are identifier
		// characters there even for a candidate containing neither K nor S.
		if unusedPrivateConstASCIIIdentifier(name) && !hasUnicodeASCIIFold {
			index.byName[name] = unusedPrivateConstNonOverlapping(index.byName[name])
			continue
		}
		index.byName[name] = nil
		matcher := unusedNameMatcher(name)
		for lineNumber, line := range lines {
			index.stats.fallbackRegexLineScan++
			for _, location := range matcher.FindAllStringIndex(line, -1) {
				begin := location[0]
				if !isIdentifierByte(line[begin]) {
					begin++
				}
				index.byName[name] = append(index.byName[name], unusedPrivateConstOccurrence{
					line: lineNumber + 1, begin: begin, matchFrom: location[0], matchTo: location[1],
				})
			}
		}
	}
	return index
}

func unusedPrivateConstTokenOccurrence(line string, lineNumber, start, end int) unusedPrivateConstOccurrence {
	matchFrom := start
	begin := start
	if start > 0 {
		_, size := utf8.DecodeLastRuneInString(line[:start])
		matchFrom -= size
		begin = matchFrom
		if !isIdentifierByte(line[matchFrom]) {
			begin++
		}
	}
	matchTo := end
	if end < len(line) {
		_, size := utf8.DecodeRuneInString(line[end:])
		matchTo += size
	}
	return unusedPrivateConstOccurrence{line: lineNumber, begin: begin, matchFrom: matchFrom, matchTo: matchTo}
}

func unusedPrivateConstNonOverlapping(occurrences []unusedPrivateConstOccurrence) []unusedPrivateConstOccurrence {
	if len(occurrences) < 2 {
		return occurrences
	}
	filtered := occurrences[:0]
	previousLine, previousEnd := 0, 0
	for _, occurrence := range occurrences {
		if occurrence.line == previousLine && occurrence.matchFrom < previousEnd {
			continue
		}
		filtered = append(filtered, occurrence)
		previousLine, previousEnd = occurrence.line, occurrence.matchTo
	}
	return filtered
}

func unusedPrivateConstASCIIIdentifier(name string) bool {
	if name == "" {
		return false
	}
	for i := 0; i < len(name); i++ {
		if !unusedPrivateConstASCIIIdentifierByte(name[i]) {
			return false
		}
	}
	return true
}

func unusedPrivateConstASCIIIdentifierByte(b byte) bool {
	return b == '_' || b >= '0' && b <= '9' || b >= 'a' && b <= 'z' || b >= 'A' && b <= 'Z'
}

func unusedPrivateConstASCIIIdentifierHash(text string) uint64 {
	hash := uint64(14695981039346656037)
	for i := 0; i < len(text); i++ {
		b := text[i]
		if b >= 'A' && b <= 'Z' {
			b += 'a' - 'A'
		}
		hash ^= uint64(b)
		hash *= 1099511628211
	}
	return hash
}

func unusedPrivateConstASCIIEqualFold(left, right string) bool {
	if len(left) != len(right) {
		return false
	}
	for i := 0; i < len(left); i++ {
		leftByte, rightByte := left[i], right[i]
		if leftByte >= 'A' && leftByte <= 'Z' {
			leftByte += 'a' - 'A'
		}
		if rightByte >= 'A' && rightByte <= 'Z' {
			rightByte += 'a' - 'A'
		}
		if leftByte != rightByte {
			return false
		}
	}
	return true
}

var unusedPrivateConstUnicodeFoldASCII = func() map[rune]byte {
	folds := make(map[rune]byte)
	for ascii := rune('A'); ascii <= 'Z'; ascii++ {
		for folded := unicode.SimpleFold(ascii); folded != ascii; folded = unicode.SimpleFold(folded) {
			if folded >= utf8.RuneSelf {
				folds[folded] = byte(ascii + ('a' - 'A'))
			}
		}
	}
	return folds
}()

// qualifiedModuleReference reports whether the text immediately before a name
// occurrence is `<module>.`, which binds to the module-level member even when
// a same-named local shadows unqualified references.
func qualifiedModuleReference(prefix, moduleName string) bool {
	moduleName = strings.TrimSpace(moduleName)
	if moduleName == "" {
		return false
	}
	dot := strings.LastIndexByte(prefix, '.')
	if dot < 0 {
		return false
	}
	qualifier := strings.TrimSpace(prefix[:dot])
	end := len(qualifier)
	start := end
	for start > 0 && isIdentifierByte(qualifier[start-1]) {
		start--
	}
	if !strings.EqualFold(qualifier[start:end], moduleName) {
		return false
	}
	// The qualifier itself must be a complete token: `XModule.` or `A.B.`
	// before the name does not denote this module's member.
	return start == 0 || (!isIdentifierByte(qualifier[start-1]) && qualifier[start-1] != '.')
}
