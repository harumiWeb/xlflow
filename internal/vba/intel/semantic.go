package intel

import (
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode"
)

const (
	SemanticTokenNamespace  = "namespace"
	SemanticTokenType       = "type"
	SemanticTokenClass      = "class"
	SemanticTokenEnum       = "enum"
	SemanticTokenInterface  = "interface"
	SemanticTokenParameter  = "parameter"
	SemanticTokenVariable   = "variable"
	SemanticTokenProperty   = "property"
	SemanticTokenEnumMember = "enumMember"
	SemanticTokenEvent      = "event"
	SemanticTokenFunction   = "function"
	SemanticTokenMethod     = "method"
	SemanticTokenKeyword    = "keyword"
	SemanticTokenModifier   = "modifier"
	SemanticTokenComment    = "comment"
	SemanticTokenString     = "string"
	SemanticTokenNumber     = "number"
	SemanticTokenOperator   = "operator"
)

const (
	SemanticModifierDeclaration    = "declaration"
	SemanticModifierDefinition     = "definition"
	SemanticModifierReadonly       = "readonly"
	SemanticModifierStatic         = "static"
	SemanticModifierDefaultLibrary = "defaultLibrary"
)

var SemanticTokenTypes = []string{
	SemanticTokenNamespace,
	SemanticTokenType,
	SemanticTokenClass,
	SemanticTokenEnum,
	SemanticTokenInterface,
	SemanticTokenParameter,
	SemanticTokenVariable,
	SemanticTokenProperty,
	SemanticTokenEnumMember,
	SemanticTokenEvent,
	SemanticTokenFunction,
	SemanticTokenMethod,
	SemanticTokenKeyword,
	SemanticTokenModifier,
	SemanticTokenComment,
	SemanticTokenString,
	SemanticTokenNumber,
	SemanticTokenOperator,
}

var SemanticTokenModifiers = []string{
	SemanticModifierDeclaration,
	SemanticModifierDefinition,
	SemanticModifierReadonly,
	SemanticModifierStatic,
	SemanticModifierDefaultLibrary,
}

type SemanticToken struct {
	Range     Range
	Type      string
	Modifiers []string
}

type semanticBuilder struct {
	analyzer Analyzer
	doc      Document
	tokens   []SemanticToken
}

var semanticKeywords = map[string]bool{
	"if": true, "then": true, "else": true, "elseif": true, "end": true, "select": true, "case": true,
	"for": true, "each": true, "in": true, "to": true, "step": true, "next": true, "do": true, "loop": true,
	"while": true, "wend": true, "until": true, "with": true, "on": true, "error": true, "resume": true,
	"goto": true, "gosub": true, "return": true, "exit": true, "call": true, "set": true, "let": true,
	"get": true, "raiseevent": true, "load": true, "unload": true, "redim": true, "preserve": true,
	"erase": true, "open": true, "close": true, "input": true, "print": true, "write": true, "lock": true,
	"unlock": true, "name": true, "kill": true, "mkdir": true, "rmdir": true, "chdir": true, "chdrive": true,
	"option": true, "explicit": true, "base": true, "compare": true, "private": true, "module": true,
	"sub": true, "function": true, "property": true, "type": true, "enum": true, "const": true,
	"declare": true, "lib": true, "alias": true, "implements": true, "as": true, "new": true,
	"true": true, "false": true, "nothing": true, "null": true, "empty": true, "missing": true,
}

var semanticModifiers = map[string]bool{
	"public": true, "private": true, "friend": true, "static": true, "byval": true, "byref": true,
	"optional": true, "paramarray": true, "global": true, "ptrsafe": true, "withevents": true,
}

var operatorRunes = map[rune]bool{
	'=': true, '<': true, '>': true, '+': true, '-': true, '*': true, '/': true, '\\': true,
	'^': true, '&': true, ':': true,
}

var memberExprRe = regexp.MustCompile(`(?i)([A-Za-z_][A-Za-z0-9_]*(?:\s*\([^()\r\n]*\))?(?:\.[A-Za-z_][A-Za-z0-9_]*(?:\s*\([^()\r\n]*\))?)*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)`)
var procedureDeclRe = regexp.MustCompile(`(?i)\b(?:Public|Private|Friend|Static)?\s*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)`)
var propertyDeclRe = regexp.MustCompile(`(?i)\b(?:Public|Private|Friend|Static)?\s*(Property)\s+(?:Get|Let|Set)\s+([A-Za-z_][A-Za-z0-9_]*)`)
var projectTypeReferenceRe = regexp.MustCompile(`(?i)\b(?:As|New|Implements)\s+(?:[A-Za-z_][A-Za-z0-9_]*\s*\.\s*)*([A-Za-z_][A-Za-z0-9_]*)`)

func (a Analyzer) SemanticTokens(doc Document, open []Document) ([]SemanticToken, error) {
	builder := semanticBuilder{analyzer: a, doc: doc}
	builder.addLexicalTokens()
	builder.addSymbolTokens(open)
	builder.addParameterReferenceTokens()
	builder.addProjectTypeReferenceTokens(open)
	builder.addProjectProcedureReferenceTokens(open)
	builder.addKnownIdentifierTokens(open)
	builder.addMemberTokens()
	builder.addUserFormControlTokens()
	return normalizeSemanticTokens(builder.tokens), nil
}

// addParameterReferenceTokens applies the parameter classification to uses in
// the declaring procedure as well as to the declaration itself. TextMate can
// recognize only the declaration; these tokens keep the two appearances
// visually consistent without confusing members or named arguments.
func (b *semanticBuilder) addParameterReferenceTokens() {
	syms, err := b.analyzer.DocumentSymbols(b.doc)
	if err != nil {
		return
	}
	for _, sym := range syms {
		if !strings.EqualFold(sym.Kind, "parameter") || sym.Name == "" {
			continue
		}
		scope, ok := currentProcedureRangeForDocument(b.doc, sym.Selection.Start)
		if !ok {
			continue
		}
		for _, rng := range codeIdentifierRanges(b.doc.Source, sym.Name) {
			if rangeContains(scope, rng) && b.isParameterReference(rng) {
				b.add(rng, SemanticTokenParameter)
			}
		}
	}
}

func (b *semanticBuilder) isParameterReference(rng Range) bool {
	line := lineAt(b.doc.Source, rng.Start.Line)
	start := byteIndexForUTF16(line, rng.Start.Character)
	end := byteIndexForUTF16(line, rng.End.Character)
	if start > 0 && line[start-1] == '.' {
		return false
	}
	return !strings.HasPrefix(strings.TrimSpace(line[end:]), ":=")
}

// addProjectTypeReferenceTokens highlights type positions such as "As Customer"
// and "New Invoice" for types declared in the current workspace. Restricting
// this to explicit type contexts avoids mistaking an ordinary variable that
// happens to share a name with a type for a type reference.
func (b *semanticBuilder) addProjectTypeReferenceTokens(open []Document) {
	syms, err := b.analyzer.WorkspaceSymbols(open, "")
	if err != nil {
		return
	}
	types := make(map[string]string)
	for _, sym := range syms {
		tokenType := semanticTypeForProjectTypeSymbol(sym)
		if tokenType == "" || !b.analyzer.visibleCompletionSymbol(b.doc, "", sym) {
			continue
		}
		types[strings.ToLower(sym.Name)] = tokenType
	}
	for lineNo, line := range normalizedLines(b.doc.Source) {
		limit := codeLimit(line)
		for _, match := range projectTypeReferenceRe.FindAllStringSubmatchIndex(line[:limit], -1) {
			if len(match) < 4 || match[2] < 0 || match[3] < 0 {
				continue
			}
			name := strings.ToLower(line[match[2]:match[3]])
			if tokenType, ok := types[name]; ok {
				b.add(byteRange(lineNo, line, match[2], match[3]), tokenType)
			}
		}
	}
}

func semanticTypeForProjectTypeSymbol(sym Symbol) string {
	switch strings.ToLower(sym.Kind) {
	case "type":
		return SemanticTokenType
	case "enum":
		return SemanticTokenEnum
	case "class":
		return SemanticTokenClass
	default:
		return ""
	}
}

// addProjectProcedureReferenceTokens classifies calls to workspace Sub and
// Function declarations. Declarations are already covered by addSymbolTokens;
// this supplies the same function color for bare Sub calls and Functions used
// in expressions.
func (b *semanticBuilder) addProjectProcedureReferenceTokens(open []Document) {
	syms, err := b.analyzer.WorkspaceSymbols(open, "")
	if err != nil {
		return
	}
	procedures := make(map[string]bool)
	for _, sym := range syms {
		if !isProjectProcedureSymbol(sym) || !b.analyzer.visibleCompletionSymbol(b.doc, "", sym) {
			continue
		}
		procedures[strings.ToLower(sym.Name)] = true
	}
	if len(procedures) == 0 {
		return
	}

	docSyms, err := b.analyzer.DocumentSymbols(b.doc)
	if err != nil {
		return
	}
	declarations := make(map[string]bool)
	locals := make(map[string]bool)
	for _, sym := range docSyms {
		if isProjectProcedureSymbol(sym) {
			declarations[semanticRangeKey(b.symbolNameRange(sym))] = true
		}
		if isLocalSymbol(sym) {
			locals[procedureLocalKey(sym.Parent, sym.Name)] = true
		}
	}
	for lineNo, line := range normalizedLines(b.doc.Source) {
		for _, span := range codeIdentifierSpans(line) {
			name := line[span.start:span.end]
			if !procedures[strings.ToLower(name)] {
				continue
			}
			rng := byteRange(lineNo, line, span.start, span.end)
			if declarations[semanticRangeKey(rng)] || !b.isProjectProcedureReference(rng) {
				continue
			}
			procedure := currentProcedureNameForDocument(b.doc, rng.Start)
			if locals[procedureLocalKey(procedure, name)] {
				continue
			}
			b.add(rng, SemanticTokenFunction)
		}
	}
}

func isProjectProcedureSymbol(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "sub", "function", "declare_sub", "declare_function":
		return true
	default:
		return false
	}
}

func procedureLocalKey(procedure, name string) string {
	return strings.ToLower(procedure) + "\x00" + strings.ToLower(name)
}

func (b *semanticBuilder) isProjectProcedureReference(rng Range) bool {
	line := lineAt(b.doc.Source, rng.Start.Line)
	start := byteIndexForUTF16(line, rng.Start.Character)
	end := byteIndexForUTF16(line, rng.End.Character)
	before := strings.TrimSpace(line[:start])
	if strings.HasSuffix(before, ".") {
		return false
	}
	return !strings.HasPrefix(strings.TrimSpace(line[end:]), ":=")
}

func (b *semanticBuilder) addLexicalTokens() {
	for lineNo, line := range normalizedLines(b.doc.Source) {
		limit := len(line)
		for i := 0; i < limit; {
			r, size := firstRune(line[i:])
			if r == '\'' {
				b.addBytes(lineNo, line, i, len(line), SemanticTokenComment)
				break
			}
			if r == '"' {
				end := scanStringEnd(line, i+size)
				b.addBytes(lineNo, line, i, end, SemanticTokenString)
				i = end
				continue
			}
			if r == '#' {
				end := scanDateOrDirectiveEnd(line, i+size)
				tokType := SemanticTokenNumber
				if end == i+size || strings.HasPrefix(strings.TrimSpace(line[i:]), "#") && strings.Contains(line[i:end], " ") {
					tokType = SemanticTokenKeyword
				}
				b.addBytes(lineNo, line, i, end, tokType)
				i = end
				continue
			}
			if unicode.IsDigit(r) || (r == '.' && i+size < limit && nextRuneIsDigit(line[i+size:])) {
				end := scanNumberEnd(line, i+size)
				b.addBytes(lineNo, line, i, end, SemanticTokenNumber)
				i = end
				continue
			}
			if operatorRunes[r] {
				b.addBytes(lineNo, line, i, i+size, SemanticTokenOperator)
				i += size
				continue
			}
			if isIdentStartRune(r) {
				end := i + size
				for end < limit {
					next, nextSize := firstRune(line[end:])
					if !isIdentRune(next) {
						break
					}
					end += nextSize
				}
				word := strings.ToLower(line[i:end])
				switch {
				case semanticModifiers[word]:
					b.addBytes(lineNo, line, i, end, SemanticTokenModifier)
				case semanticKeywords[word]:
					b.addBytes(lineNo, line, i, end, SemanticTokenKeyword)
				case word == "and" || word == "or" || word == "not" || word == "xor" || word == "eqv" || word == "imp" || word == "is" || word == "like" || word == "mod":
					b.addBytes(lineNo, line, i, end, SemanticTokenOperator)
				}
				i = end
				continue
			}
			i += max(1, size)
		}
	}
}

func (b *semanticBuilder) addSymbolTokens(open []Document) {
	syms, err := b.analyzer.DocumentSymbols(b.doc)
	if err != nil {
		return
	}
	for _, sym := range syms {
		if sym.Name == "" || sym.Selection == (Range{}) {
			continue
		}
		tokenType := semanticTypeForSymbol(sym)
		if tokenType == "" {
			continue
		}
		rng := b.symbolNameRange(sym)
		mods := []string{SemanticModifierDeclaration, SemanticModifierDefinition}
		if strings.EqualFold(sym.Visibility, "Static") || strings.Contains(strings.ToLower(sym.Detail), "static ") {
			mods = append(mods, SemanticModifierStatic)
		}
		if strings.EqualFold(sym.Kind, "const") {
			mods = append(mods, SemanticModifierReadonly)
		}
		b.add(rng, tokenType, mods...)
	}
	if all, err := b.analyzer.WorkspaceSymbols(open, ""); err == nil {
		for _, sym := range all {
			if !sameDocumentSymbol(sym, b.doc) {
				continue
			}
			tokenType := semanticTypeForSymbol(sym)
			if tokenType != "" {
				b.add(b.symbolNameRange(sym), tokenType, SemanticModifierDeclaration, SemanticModifierDefinition)
			}
		}
	}
	b.addProcedureDeclarationFallbackTokens()
}

func (b *semanticBuilder) addKnownIdentifierTokens(open []Document) {
	_ = open
	for lineNo, line := range normalizedLines(b.doc.Source) {
		limit := codeLimit(line)
		for start := 0; start < limit; {
			r, size := firstRune(line[start:limit])
			if !isIdentStartRune(r) {
				start += max(1, size)
				continue
			}
			end := start + size
			for end < limit {
				r, size = firstRune(line[end:limit])
				if !isIdentRune(r) {
					break
				}
				end += size
			}
			word := line[start:end]
			lower := strings.ToLower(word)
			if semanticKeywords[lower] || semanticModifiers[lower] {
				start = end
				continue
			}
			rng := byteRange(lineNo, line, start, end)
			switch {
			case b.analyzer.DB != nil:
				if typ, ok := b.analyzer.DB.ResolveType(word); ok {
					b.add(rng, semanticTypeForDBType(typ.Kind), SemanticModifierDefaultLibrary)
				} else if c, ok := b.analyzer.DB.ResolveConstant(word); ok {
					tokenType := SemanticTokenVariable
					if c.EnumGroup != "" {
						tokenType = SemanticTokenEnumMember
					}
					b.add(rng, tokenType, SemanticModifierReadonly, SemanticModifierDefaultLibrary)
				} else if _, ok := b.analyzer.DB.ResolveGlobal(word); ok {
					b.add(rng, SemanticTokenVariable, SemanticModifierDefaultLibrary)
				} else if m, ok := b.analyzer.DB.ResolveMember("VBA.Global", word); ok {
					kind := SemanticTokenFunction
					if len(m.Parameters) == 0 && m.ReturnType != "" {
						kind = SemanticTokenProperty
					}
					b.add(rng, kind, SemanticModifierDefaultLibrary)
				}
			}
			start = end
		}
	}
}

func (b *semanticBuilder) addMemberTokens() {
	if b.analyzer.DB == nil {
		return
	}
	for lineNo, line := range normalizedLines(b.doc.Source) {
		limit := codeLimit(line)
		for _, match := range memberExprRe.FindAllStringSubmatchIndex(line[:limit], -1) {
			if len(match) < 6 || match[2] < 0 || match[3] < 0 || match[4] < 0 || match[5] < 0 {
				continue
			}
			receiverExpr := strings.TrimSpace(line[match[2]:match[3]])
			memberName := line[match[4]:match[5]]
			pos := Position{Line: lineNo, Character: utf16Len(line[:match[4]])}
			offset := byteOffsetForPosition(b.doc.Source, pos)
			receiverType, ok := b.analyzer.resolveDocumentExpressionTypeAt(b.doc, receiverExpr, offset)
			if !ok {
				if strings.HasPrefix(strings.TrimSpace(receiverExpr), ".") {
					receiverType, ok = b.analyzer.withBlockTypeAt(b.doc, pos, offset)
				}
			}
			if !ok {
				continue
			}
			member, ok := b.analyzer.DB.ResolveMember(receiverType, memberName)
			if !ok {
				continue
			}
			tokenType := SemanticTokenProperty
			switch b.analyzer.memberKind(receiverType, memberName) {
			case "method":
				tokenType = SemanticTokenMethod
			case "event":
				tokenType = SemanticTokenEvent
			}
			mods := []string{SemanticModifierDefaultLibrary}
			if member.ReadOnly {
				mods = append(mods, SemanticModifierReadonly)
			}
			b.add(byteRange(lineNo, line, match[4], match[5]), tokenType, mods...)
		}
	}
}

func (b *semanticBuilder) addProcedureDeclarationFallbackTokens() {
	for lineNo, line := range normalizedLines(b.doc.Source) {
		limit := codeLimit(line)
		for _, match := range procedureDeclRe.FindAllStringSubmatchIndex(line[:limit], -1) {
			if len(match) >= 6 && match[4] >= 0 && match[5] >= 0 {
				b.add(byteRange(lineNo, line, match[4], match[5]), SemanticTokenFunction, SemanticModifierDeclaration, SemanticModifierDefinition)
			}
		}
		for _, match := range propertyDeclRe.FindAllStringSubmatchIndex(line[:limit], -1) {
			if len(match) >= 6 && match[4] >= 0 && match[5] >= 0 {
				b.add(byteRange(lineNo, line, match[4], match[5]), SemanticTokenProperty, SemanticModifierDeclaration, SemanticModifierDefinition)
			}
		}
	}
}

func (b *semanticBuilder) symbolNameRange(sym Symbol) Range {
	lineNo := sym.Selection.Start.Line
	if lineNo < 0 {
		return sym.Selection
	}
	line := lineAt(b.doc.Source, lineNo)
	if line == "" || sym.Name == "" {
		return sym.Selection
	}
	startByte := byteIndexForUTF16(line, sym.Selection.Start.Character)
	if startByte > len(line) {
		startByte = len(line)
	}
	lowerLine := strings.ToLower(line)
	lowerName := strings.ToLower(sym.Name)
	if idx := strings.Index(lowerLine[startByte:], lowerName); idx >= 0 {
		start := startByte + idx
		end := start + len(sym.Name)
		if tokenBoundary(line, start, end) {
			return byteRange(lineNo, line, start, end)
		}
	}
	if idx := strings.Index(lowerLine, lowerName); idx >= 0 {
		end := idx + len(sym.Name)
		if tokenBoundary(line, idx, end) {
			return byteRange(lineNo, line, idx, end)
		}
	}
	return sym.Selection
}

func (b *semanticBuilder) addUserFormControlTokens() {
	if !b.analyzer.isFormDocument(b.doc) {
		return
	}
	for _, control := range b.analyzer.formControls(b.doc) {
		for _, rng := range identifierRanges(b.doc.Source, control.Name) {
			b.add(rng, SemanticTokenVariable)
		}
	}
}

func (b *semanticBuilder) addBytes(lineNo int, line string, start, end int, tokenType string, mods ...string) {
	b.add(byteRange(lineNo, line, start, end), tokenType, mods...)
}

func (b *semanticBuilder) add(r Range, tokenType string, mods ...string) {
	if tokenType == "" || r.End.Line != r.Start.Line || r.End.Character <= r.Start.Character {
		return
	}
	b.tokens = append(b.tokens, SemanticToken{Range: r, Type: tokenType, Modifiers: mods})
}

func semanticTypeForSymbol(sym Symbol) string {
	switch strings.ToLower(sym.Kind) {
	case "module":
		return SemanticTokenNamespace
	case "class":
		return SemanticTokenClass
	case "type":
		return SemanticTokenType
	case "enum":
		return SemanticTokenEnum
	case "sub", "function", "declare_sub", "declare_function":
		return SemanticTokenFunction
	case "property", "property_get", "property_let", "property_set":
		return SemanticTokenProperty
	case "const":
		return SemanticTokenVariable
	case "field", "module_variable", "local_variable":
		return SemanticTokenVariable
	case "parameter":
		return SemanticTokenParameter
	case "implements":
		return SemanticTokenInterface
	case "event":
		return SemanticTokenEvent
	default:
		return ""
	}
}

func semanticTypeForDBType(kind string) string {
	switch strings.ToLower(kind) {
	case "class", "collection":
		return SemanticTokenClass
	case "enum":
		return SemanticTokenEnum
	case "interface":
		return SemanticTokenInterface
	default:
		return SemanticTokenType
	}
}

func normalizeSemanticTokens(tokens []SemanticToken) []SemanticToken {
	sort.SliceStable(tokens, func(i, j int) bool {
		a, b := tokens[i].Range, tokens[j].Range
		if a.Start.Line != b.Start.Line {
			return a.Start.Line < b.Start.Line
		}
		if a.Start.Character != b.Start.Character {
			return a.Start.Character < b.Start.Character
		}
		if a.End.Character != b.End.Character {
			return a.End.Character < b.End.Character
		}
		return semanticPriority(tokens[i].Type) < semanticPriority(tokens[j].Type)
	})
	out := tokens[:0]
	seen := map[string]bool{}
	lastLine, lastEnd := -1, -1
	for _, tok := range tokens {
		if tok.Range.End.Line != tok.Range.Start.Line || tok.Range.End.Character <= tok.Range.Start.Character {
			continue
		}
		key := semanticRangeKey(tok.Range)
		if seen[key] {
			continue
		}
		if tok.Range.Start.Line == lastLine && tok.Range.Start.Character < lastEnd {
			continue
		}
		seen[key] = true
		out = append(out, tok)
		lastLine = tok.Range.Start.Line
		lastEnd = tok.Range.End.Character
	}
	return out
}

func semanticPriority(tokenType string) int {
	switch tokenType {
	case SemanticTokenKeyword, SemanticTokenModifier, SemanticTokenString, SemanticTokenNumber, SemanticTokenComment, SemanticTokenOperator:
		return 10
	default:
		return 0
	}
}

func semanticRangeKey(r Range) string {
	return strings.Join([]string{
		intString(r.Start.Line),
		intString(r.Start.Character),
		intString(r.End.Line),
		intString(r.End.Character),
	}, ":")
}

func sameDocumentSymbol(sym Symbol, doc Document) bool {
	target := firstNonEmpty(doc.URI, doc.Path)
	if target == "" {
		return false
	}
	return strings.EqualFold(sym.File, target) || strings.EqualFold(pathKey(sym.File), pathKey(doc.Path))
}

func byteRange(lineNo int, line string, start, end int) Range {
	start = max(0, min(start, len(line)))
	end = max(start, min(end, len(line)))
	return Range{
		Start: Position{Line: lineNo, Character: utf16Len(line[:start])},
		End:   Position{Line: lineNo, Character: utf16Len(line[:end])},
	}
}

func scanStringEnd(line string, start int) int {
	for i := start; i < len(line); i++ {
		if line[i] == '"' {
			if i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			return i + 1
		}
	}
	return len(line)
}

func scanDateOrDirectiveEnd(line string, start int) int {
	for i := start; i < len(line); i++ {
		if line[i] == '#' {
			return i + 1
		}
		if unicode.IsSpace(rune(line[i])) && start == 1 {
			return i
		}
	}
	return len(line)
}

func scanNumberEnd(line string, start int) int {
	i := start
	for i < len(line) {
		r, size := firstRune(line[i:])
		if !unicode.IsDigit(r) && !unicode.IsLetter(r) && r != '.' && r != '&' && r != '_' && r != '+' && r != '-' {
			break
		}
		i += max(1, size)
	}
	return i
}

func isDigitRune(r rune) bool {
	return r >= '0' && r <= '9'
}

func intString(value int) string {
	return strconv.Itoa(value)
}

func nextRuneIsDigit(s string) bool {
	if s == "" {
		return false
	}
	r, _ := firstRune(s)
	return isDigitRune(r)
}

func tokenBoundary(line string, start, end int) bool {
	if start > 0 {
		r, _ := lastRune(line[:start])
		if isIdentRune(r) {
			return false
		}
	}
	if end < len(line) {
		r, _ := firstRune(line[end:])
		if isIdentRune(r) {
			return false
		}
	}
	return true
}
