package intel

import (
	"errors"
	"fmt"
	"regexp"
	"sort"
	"strings"
)

type RenameTarget struct {
	Name   string
	Kind   string
	Range  Range
	Symbol Symbol
	Reason string
}

type RenameEdit struct {
	URI     string
	Path    string
	Range   Range
	NewText string
}

type symbolIdentity struct {
	File       string
	Module     string
	ModuleKind string
	Parent     string
	Kind       string
	Range      Range
	Name       string
}

func (a Analyzer) PrepareRename(doc Document, pos Position, open []Document) (*RenameTarget, error) {
	target, err := a.renameTarget(doc, pos, open)
	if err != nil {
		return nil, err
	}
	return &target, nil
}

func (a Analyzer) Rename(doc Document, pos Position, newName string, open []Document, uriForPath func(string) string) ([]RenameEdit, error) {
	if err := validateRenameName(newName); err != nil {
		return nil, err
	}
	target, err := a.renameTarget(doc, pos, open)
	if err != nil {
		return nil, err
	}
	var ranges []Range
	if isRenameLabel(target.Symbol) {
		ranges = labelRenameRanges(doc, target.Symbol)
	} else {
		ranges = a.symbolRenameRanges(doc, target.Symbol, open)
	}
	if len(ranges) == 0 {
		return nil, fmt.Errorf("cannot rename unresolved identifier")
	}
	sort.SliceStable(ranges, func(i, j int) bool {
		if ranges[i].Start.Line != ranges[j].Start.Line {
			return ranges[i].Start.Line < ranges[j].Start.Line
		}
		return ranges[i].Start.Character < ranges[j].Start.Character
	})
	uri := doc.URI
	if uri == "" && uriForPath != nil {
		uri = uriForPath(doc.Path)
	}
	edits := make([]RenameEdit, 0, len(ranges))
	seen := map[string]bool{}
	for _, r := range ranges {
		key := locationKey(doc.Path, r)
		if seen[key] {
			continue
		}
		seen[key] = true
		edits = append(edits, RenameEdit{URI: uri, Path: doc.Path, Range: r, NewText: newName})
	}
	return edits, nil
}

func (a Analyzer) renameTarget(doc Document, pos Position, open []Document) (RenameTarget, error) {
	word, wordRange := WordAt(doc.Source, pos)
	if strings.TrimSpace(word) == "" || !rangeContains(wordRange, Range{Start: pos, End: pos}) {
		return RenameTarget{}, fmt.Errorf("cannot rename unresolved identifier")
	}
	if !rangeIsCodeIdentifier(doc.Source, wordRange, word) {
		return RenameTarget{}, fmt.Errorf("cannot rename unresolved identifier")
	}
	if a.isExternalRenameTarget(doc, word, wordRange, pos) {
		return RenameTarget{}, fmt.Errorf("cannot rename external host member")
	}
	sym, err := a.resolveRenameSymbol(doc, pos, open, word, wordRange)
	if err != nil {
		return RenameTarget{}, err
	}
	if reason := renameUnsupportedReason(sym); reason != "" {
		return RenameTarget{}, errors.New(reason)
	}
	return RenameTarget{Name: sym.Name, Kind: sym.Kind, Range: sym.Selection, Symbol: sym}, nil
}

func (a Analyzer) resolveRenameSymbol(doc Document, pos Position, open []Document, word string, wordRange Range) (Symbol, error) {
	docSyms, err := a.DocumentSymbols(doc)
	if err != nil {
		return Symbol{}, err
	}
	for _, sym := range docSyms {
		if strings.EqualFold(sym.Name, word) && rangeContains(sym.Selection, wordRange) {
			return sym, nil
		}
	}
	defs, err := a.definitionSymbols(doc, pos, open, word)
	if err != nil {
		return Symbol{}, err
	}
	if len(defs) > 1 {
		return Symbol{}, fmt.Errorf("cannot rename ambiguous symbol")
	}
	var supported []Symbol
	for _, sym := range defs {
		if renameUnsupportedReason(sym) == "" {
			supported = append(supported, sym)
		}
	}
	if len(supported) == 0 {
		if len(defs) > 0 {
			return Symbol{}, errors.New(renameUnsupportedReason(defs[0]))
		}
		return Symbol{}, fmt.Errorf("cannot rename unresolved identifier")
	}
	if len(supported) != 1 {
		return Symbol{}, fmt.Errorf("cannot rename ambiguous symbol")
	}
	return supported[0], nil
}

func (a Analyzer) symbolRenameRanges(doc Document, target Symbol, open []Document) []Range {
	word := target.Name
	search := Range{Start: Position{Line: 0, Character: 0}, End: Position{Line: len(normalizedLines(doc.Source)), Character: 0}}
	if isLocalSymbol(target) {
		if scope, ok := currentProcedureRangeForDocument(doc, target.Selection.Start); ok {
			search = scope
		}
	}
	targetID := renameSymbolIdentity(target)
	var out []Range
	for _, r := range codeIdentifierRanges(doc.Source, word) {
		if !rangeContains(search, r) {
			continue
		}
		pos := r.Start
		pos.Character++
		sym, err := a.resolveRenameSymbol(doc, pos, open, word, r)
		if err != nil {
			continue
		}
		if renameSymbolIdentity(sym) == targetID {
			out = append(out, r)
		}
	}
	return out
}

func renameSymbolIdentity(sym Symbol) symbolIdentity {
	return symbolIdentity{
		File:       pathKey(sym.File),
		Module:     strings.ToLower(sym.Module),
		ModuleKind: strings.ToLower(sym.ModuleKind),
		Parent:     strings.ToLower(sym.Parent),
		Kind:       strings.ToLower(sym.Kind),
		Range:      sym.Selection,
		Name:       strings.ToLower(sym.Name),
	}
}

func renameUnsupportedReason(sym Symbol) string {
	switch strings.ToLower(sym.Kind) {
	case "local_variable", "parameter":
		return ""
	case "const":
		if sym.Parent != "" {
			return ""
		}
		if isPrivateOrImplicit(sym) {
			return ""
		}
		return "project-wide public rename is not supported yet"
	case "module_variable":
		if isPrivateOrImplicit(sym) {
			return ""
		}
		return "project-wide public rename is not supported yet"
	case "sub", "function":
		if isUserFormEventProcedure(sym) {
			return "userform control/event rename is not supported yet"
		}
		if sym.Parent == "" && strings.EqualFold(sym.Visibility, "Private") {
			return ""
		}
		return "project-wide public rename is not supported yet"
	case "label":
		return ""
	case "field", "withevents_field":
		return "userform control/event rename is not supported yet"
	case "module", "class":
		return "cannot rename module files"
	case "property", "property_get", "property_let", "property_set":
		return "property group rename is not supported yet"
	default:
		return "cannot rename ambiguous symbol"
	}
}

func isPrivateOrImplicit(sym Symbol) bool {
	return sym.Visibility == "" || strings.EqualFold(sym.Visibility, "Private")
}

func isRenameLabel(sym Symbol) bool {
	return strings.EqualFold(sym.Kind, "label")
}

func (a Analyzer) isExternalRenameTarget(doc Document, word string, wordRange Range, pos Position) bool {
	if a.DB == nil {
		return false
	}
	line := lineAt(doc.Source, wordRange.Start.Line)
	startByte := byteIndexForUTF16(line, wordRange.Start.Character)
	if startByte > len(line) {
		startByte = len(line)
	}
	beforeWord := strings.TrimRight(line[:startByte], " \t")
	if strings.HasSuffix(beforeWord, ".") {
		receiverExpr := expressionBefore(strings.TrimSuffix(beforeWord, "."))
		offset := byteOffsetForPosition(doc.Source, pos)
		if receiverExpr == "" {
			if receiverType, ok := a.withBlockTypeAt(doc, pos, offset); ok {
				if _, ok := a.DB.ResolveMember(receiverType, word); ok {
					return true
				}
			}
		} else if receiverType, ok := a.resolveDocumentExpressionTypeAt(doc, receiverExpr, offset); ok {
			if _, ok := a.DB.ResolveMember(receiverType, word); ok {
				return true
			}
		}
	}
	if _, ok := a.DB.ResolveType(word); ok {
		return true
	}
	if _, ok := a.DB.ResolveConstant(word); ok {
		return true
	}
	if _, ok := a.DB.ResolveGlobal(word); ok {
		return true
	}
	if _, ok := a.DB.ResolveMember("Excel.Application", word); ok {
		return true
	}
	if _, ok := a.DB.ResolveMember("VBA.Global", word); ok {
		return true
	}
	return false
}

func rangeIsCodeIdentifier(source string, r Range, word string) bool {
	for _, candidate := range codeIdentifierRanges(source, word) {
		if candidate == r {
			return true
		}
	}
	return false
}

func codeIdentifierRanges(source, name string) []Range {
	if strings.TrimSpace(name) == "" {
		return nil
	}
	var out []Range
	for lineNo, line := range normalizedLines(source) {
		for _, span := range codeIdentifierSpans(line) {
			if strings.EqualFold(line[span.start:span.end], name) {
				out = append(out, byteRange(lineNo, line, span.start, span.end))
			}
		}
	}
	return out
}

type byteSpan struct {
	start int
	end   int
}

func codeIdentifierSpans(line string) []byteSpan {
	var out []byteSpan
	inString := false
	for start := 0; start < len(line); {
		r, size := firstRune(line[start:])
		if r == '"' {
			start += size
			if inString && start < len(line) && line[start] == '"' {
				start++
				continue
			}
			inString = !inString
			continue
		}
		if r == '\'' && !inString {
			return out
		}
		if inString || !isIdentStartRune(r) {
			start += max(1, size)
			continue
		}
		end := start + size
		for end < len(line) {
			next, nextSize := firstRune(line[end:])
			if !isIdentRune(next) {
				break
			}
			end += nextSize
		}
		out = append(out, byteSpan{start: start, end: end})
		start = end
	}
	return out
}

func labelRenameRanges(doc Document, target Symbol) []Range {
	var out []Range
	if target.Selection.End.Character > target.Selection.Start.Character {
		out = append(out, target.Selection)
	}
	if target.Parent == "" {
		return out
	}
	scope, ok := currentProcedureRangeForDocument(doc, target.Selection.Start)
	if !ok {
		return out
	}
	for _, r := range labelReferenceRanges(doc.Source, target.Name) {
		if rangeContains(scope, r) {
			out = append(out, r)
		}
	}
	return out
}

func labelReferenceRanges(source, name string) []Range {
	var out []Range
	for lineNo, line := range normalizedLines(source) {
		limit := codeLimit(line)
		if limit <= 0 {
			continue
		}
		code := line[:limit]
		for _, re := range labelReferenceRegexps(name) {
			for _, match := range re.FindAllStringSubmatchIndex(code, -1) {
				if len(match) >= 4 && match[2] >= 0 && match[3] >= 0 {
					out = append(out, byteRange(lineNo, line, match[2], match[3]))
				}
			}
		}
	}
	return out
}

func labelReferenceRegexps(name string) []*regexp.Regexp {
	quoted := regexp.QuoteMeta(name)
	return []*regexp.Regexp{
		regexp.MustCompile(`(?i)\bGoTo\s+(` + quoted + `)\b`),
		regexp.MustCompile(`(?i)\bGoSub\s+(` + quoted + `)\b`),
		regexp.MustCompile(`(?i)\bResume\s+(` + quoted + `)\b`),
		regexp.MustCompile(`(?i)\bOn\s+Error\s+GoTo\s+(` + quoted + `)\b`),
	}
}

func validateRenameName(name string) error {
	if !normalVBAIdentifier(name) {
		return fmt.Errorf("invalid VBA identifier for rename")
	}
	if semanticKeywords[strings.ToLower(name)] || semanticModifiers[strings.ToLower(name)] {
		return fmt.Errorf("invalid VBA identifier for rename")
	}
	return nil
}

func normalVBAIdentifier(name string) bool {
	if name == "" {
		return false
	}
	for i, r := range name {
		if i == 0 {
			if r == '_' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' {
				continue
			}
			return false
		}
		if r == '_' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r >= '0' && r <= '9' {
			continue
		}
		return false
	}
	return true
}
