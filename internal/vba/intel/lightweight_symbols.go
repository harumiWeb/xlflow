package intel

import (
	"fmt"
	"sort"
	"strings"
)

// LightweightDocumentSymbols resolves declaration queries without forcing
// full-document symbol extraction. The boolean reports whether the query shape
// and source state were handled; callers should fall back when it is false.
func (a Analyzer) LightweightDocumentSymbols(doc Document, query WorkspaceSymbolQuery) ([]Symbol, bool) {
	if !query.Interactive && query.Mode != WorkspaceSymbolQueryExact && query.Mode != WorkspaceSymbolQueryQualified {
		return nil, false
	}
	idx, _, closeIndex, err := interactiveIndexForDocument(doc)
	defer closeIndex()
	if err != nil || idx == nil {
		return nil, false
	}
	if idx.incomplete {
		// A compact declaration parse can intentionally omit procedure bodies,
		// but recovery in the source still requires the established full-symbol
		// fallback so callers do not publish an unsafe partial result.
		return nil, false
	}
	out := idx.query(query)
	foreignQualified := false
	if query.Mode == WorkspaceSymbolQueryQualified {
		needle := strings.TrimSpace(query.Text)
		if index := strings.LastIndex(needle, "."); index >= 0 {
			qualifier := strings.TrimSpace(needle[:index])
			if qualifier == "" || !strings.EqualFold(qualifier, interactiveModuleName(idx, doc)) {
				foreignQualified = true
			}
		}
	}
	if query.Interactive {
		// The declaration index includes procedure parameters so that a local
		// exact lookup can be answered without rebuilding the document model.
		// Limit each open-document query to module declarations and the active
		// procedure's parameters; symbols from other procedures must not leak.
		out = filterInteractiveScopeSymbols(out, query, doc)
		// Parameters are part of the declaration index, so an exact query can
		// already have its local match without opening the procedure fragment.
		// Remove same-document module declarations first to preserve local
		// shadowing for hover and type inference as well as definition lookup.
		if query.Mode == WorkspaceSymbolQueryExact || query.Mode == WorkspaceSymbolQueryPrefix || query.Mode == WorkspaceSymbolQueryContains {
			out = dropShadowedInteractiveSymbols(out, filterLocalInteractiveSymbols(out, query), doc)
		}
		if !foreignQualified && interactiveQueryNeedsLocals(query, out, doc) {
			module := interactiveModuleName(idx, doc)
			locals, ok := a.localInteractiveSymbols(doc, query, module)
			if !ok {
				return nil, false
			}
			matchedLocals := filterLocalInteractiveSymbols(locals, query)
			out = dropShadowedInteractiveSymbols(out, matchedLocals, doc)
			out = append(out, matchedLocals...)
		}
		for _, control := range a.formControlSymbols(doc) {
			if interactiveSymbolMatches(control, query) {
				out = append(out, control)
			}
		}
	}
	if !query.Interactive && len(out) == 0 && !foreignQualified {
		return nil, false
	}
	return uniqueInteractiveSymbols(out), true
}

func dropShadowedInteractiveSymbols(indexed, locals []Symbol, doc Document) []Symbol {
	if len(indexed) == 0 || len(locals) == 0 {
		return indexed
	}
	shadowed := make(map[string]struct{}, len(locals))
	for _, local := range locals {
		if local.Parent != "" {
			shadowed[indexName(local.Name)] = struct{}{}
		}
	}
	if len(shadowed) == 0 {
		return indexed
	}
	out := indexed[:0]
	for _, symbol := range indexed {
		if symbol.Parent == "" && sameInteractiveDocument(symbol, doc) {
			if _, ok := shadowed[indexName(symbol.Name)]; ok {
				continue
			}
		}
		out = append(out, symbol)
	}
	return out
}

func sameInteractiveDocument(symbol Symbol, doc Document) bool {
	if symbol.File == "" || doc.Path == "" {
		return false
	}
	return sameInteractivePath(symbol.File, doc.Path)
}

func sameInteractivePath(left, right string) bool {
	if left == "" || right == "" {
		return false
	}
	if strings.EqualFold(left, right) {
		return true
	}
	leftKey, rightKey := renameFileKey(left), renameFileKey(right)
	return leftKey != "" && leftKey == rightKey
}

func interactiveModuleName(idx *interactiveDocumentIndex, doc Document) string {
	module := moduleNameForDocument(doc)
	if idx == nil {
		return module
	}
	for _, symbol := range idx.symbols {
		if strings.EqualFold(symbol.Kind, "module") && strings.TrimSpace(symbol.Module) != "" {
			return symbol.Module
		}
	}
	return module
}

func filterInteractiveScopeSymbols(symbols []Symbol, query WorkspaceSymbolQuery, doc Document) []Symbol {
	if !query.Interactive {
		return symbols
	}
	currentDocument := sameInteractivePath(query.DocumentPath, doc.Path)
	procedure := ""
	if currentDocument {
		procedure = query.Procedure
	}
	qualifiedModuleQuery := query.Mode == WorkspaceSymbolQueryQualified && strings.Contains(strings.TrimSpace(query.Text), ".")
	out := make([]Symbol, 0, len(symbols))
	for _, symbol := range symbols {
		if qualifiedModuleQuery && symbol.Parent != "" {
			continue
		}
		if symbol.Parent == "" || (procedure != "" && strings.EqualFold(symbol.Parent, procedure)) {
			out = append(out, symbol)
		}
	}
	return out
}

func interactiveQueryNeedsLocals(query WorkspaceSymbolQuery, indexed []Symbol, doc Document) bool {
	if !sameInteractivePath(query.DocumentPath, doc.Path) || query.Procedure == "" {
		return false
	}
	switch query.Mode {
	case WorkspaceSymbolQueryPrefix, WorkspaceSymbolQueryContains:
		return true
	case WorkspaceSymbolQueryExact:
		if len(indexed) == 0 {
			return true
		}
		return interactiveLocalNameDeclared(doc, query, query.Text)
	default:
		return false
	}
}

// interactiveLocalNameDeclared is a lexical shadowing check. It avoids
// parsing a large procedure when an exact query already found a module-level
// declaration, while still preserving local precedence for Dim/Static/Const
// declarations. Parameters are already indexed with their procedure parent.
func interactiveLocalNameDeclared(doc Document, request WorkspaceSymbolQuery, query string) bool {
	procedure := request.Procedure
	var scope Range
	found := false
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		for _, candidate := range snapshot.Procedures() {
			if strings.EqualFold(candidate.Name, procedure) && recoveryProcedureMatchesPosition(candidate.Range, request.RequestPosition) {
				scope, found = candidate.Range, true
				break
			}
		}
	} else {
		procedures, _ := procedureIndexForLines(documentLines(doc))
		for _, candidate := range procedures {
			if strings.EqualFold(candidate.Name, procedure) && recoveryProcedureMatchesPosition(candidate.Range, request.RequestPosition) {
				scope, found = candidate.Range, true
				break
			}
		}
	}
	if !found {
		return false
	}
	needle := indexName(query)
	if needle == "" {
		return false
	}
	lines := documentLines(doc)
	start := max(0, scope.Start.Line+1)
	end := min(len(lines), scope.End.Line)
	continuation := ""
	for lineNo := start; lineNo < end; lineNo++ {
		line := lines[lineNo]
		limit := codeLimit(line)
		if limit < 0 || limit > len(line) {
			limit = len(line)
		}
		text := strings.TrimSpace(line[:limit])
		for _, statementText := range splitRecoveryStatements(text) {
			statementText = strings.TrimSpace(statementText)
			if statementText == "" {
				continue
			}
			declaration := continuation
			if declaration == "" {
				fields := strings.Fields(statementText)
				if len(fields) == 0 {
					continue
				}
				keyword := strings.ToLower(fields[0])
				if keyword != "dim" && keyword != "static" && keyword != "const" {
					continue
				}
				declaration = strings.TrimSpace(statementText[len(fields[0]):])
			} else {
				declaration += " " + statementText
			}
			declaration = strings.TrimSpace(strings.TrimSuffix(declaration, "_"))
			if declarationNamesContain(declaration, needle) {
				return true
			}
			if strings.HasSuffix(statementText, "_") {
				continuation = declaration
			} else {
				continuation = ""
			}
		}
	}
	return false
}

func declarationNamesContain(declaration, needle string) bool {
	for _, statement := range splitRecoveryStatements(declaration) {
		fields := strings.Fields(statement)
		if len(fields) > 0 {
			switch strings.ToLower(fields[0]) {
			case "dim", "static", "const":
				statement = strings.TrimSpace(statement[len(fields[0]):])
			}
		}
		for _, part := range strings.Split(statement, ",") {
			name := strings.TrimSpace(strings.TrimSuffix(part, "_"))
			if index := strings.IndexAny(name, " (:="); index >= 0 {
				name = name[:index]
			}
			if strings.EqualFold(indexName(name), needle) {
				return true
			}
		}
	}
	return false
}

func filterLocalInteractiveSymbols(symbols []Symbol, query WorkspaceSymbolQuery) []Symbol {
	var out []Symbol
	for _, symbol := range symbols {
		if symbol.Parent == "" {
			continue
		}
		if interactiveSymbolMatches(symbol, query) {
			out = append(out, symbol)
		}
	}
	return out
}

func interactiveSymbolMatches(symbol Symbol, query WorkspaceSymbolQuery) bool {
	name := indexName(symbol.Name)
	qualified := indexName(qualifiedName(symbol.Module, symbol.Name))
	needle := indexName(query.Text)
	switch query.Mode {
	case WorkspaceSymbolQueryExact:
		return name == needle
	case WorkspaceSymbolQueryQualified:
		return qualified == needle || name == needle
	case WorkspaceSymbolQueryPrefix:
		return strings.HasPrefix(name, needle) || strings.HasPrefix(qualified, needle)
	case WorkspaceSymbolQueryContains:
		return strings.Contains(name, needle) || strings.Contains(qualified, needle)
	case WorkspaceSymbolQueryKind:
		return indexName(symbol.Kind) == needle
	case WorkspaceSymbolQueryModule:
		return indexName(symbol.Module) == needle
	default:
		return false
	}
}

func uniqueInteractiveSymbols(symbols []Symbol) []Symbol {
	seen := make(map[string]struct{}, len(symbols))
	out := make([]Symbol, 0, len(symbols))
	for _, symbol := range symbols {
		key := symbol.File + "\x00" + symbol.Name + "\x00" + symbol.Parent + "\x00" +
			formatPositionKey(symbol.Selection.Start) + "\x00" + formatPositionKey(symbol.Selection.End)
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		out = append(out, symbol)
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].File != out[j].File {
			return out[i].File < out[j].File
		}
		if out[i].Range.Start.Line != out[j].Range.Start.Line {
			return out[i].Range.Start.Line < out[j].Range.Start.Line
		}
		if out[i].Range.Start.Character != out[j].Range.Start.Character {
			return out[i].Range.Start.Character < out[j].Range.Start.Character
		}
		if out[i].Name != out[j].Name {
			return out[i].Name < out[j].Name
		}
		return out[i].Parent < out[j].Parent
	})
	return out
}

func formatPositionKey(pos Position) string {
	return fmt.Sprintf("%d:%d", pos.Line, pos.Character)
}

func (a Analyzer) localInteractiveSymbols(doc Document, request WorkspaceSymbolQuery, module string) ([]Symbol, bool) {
	procedure := request.Procedure
	var scope Range
	ok := false
	if snapshot := analysisSnapshotForDocument(doc); snapshot != nil {
		for _, candidate := range snapshot.ProcedureCatalog().Entries {
			if strings.EqualFold(candidate.Identity.CanonicalName, procedure) && recoveryProcedureMatchesPosition(candidate.Range, request.RequestPosition) {
				scope, ok = candidate.Range, true
				break
			}
		}
	} else {
		procedures, _ := procedureIndexForLines(documentLines(doc))
		for _, candidate := range procedures {
			if strings.EqualFold(candidate.Name, procedure) && recoveryProcedureMatchesPosition(candidate.Range, request.RequestPosition) {
				scope, ok = candidate.Range, true
				break
			}
		}
	}
	if !ok {
		return nil, false
	}
	start := byteOffsetForDocumentPosition(doc, scope.Start)
	end := byteOffsetForDocumentPosition(doc, scope.End)
	if end < len(doc.Source) {
		_, end = rawLineBounds(doc.Source, end)
	}
	if start < 0 || end <= start || end > len(doc.Source) {
		return nil, false
	}
	fragment := doc
	fragment.Source = doc.Source[start:end]
	fragment.Snapshot = nil
	parsed, closeParsed, err := parsedDocumentForDocument(fragment)
	if err != nil {
		return nil, false
	}
	defer closeParsed()
	symbols, err := a.inspectDocumentSourceSymbols(fragment, parsed)
	if err != nil {
		return nil, false
	}
	for i := range symbols {
		symbols[i].Range = rebaseRange(symbols[i].Range, Position{}, scope.Start)
		symbols[i].Selection = rebaseRange(symbols[i].Selection, Position{}, scope.Start)
		symbols[i].File = doc.Path
		if strings.TrimSpace(module) != "" {
			symbols[i].Module = module
		}
	}
	return symbols, true
}

func recoveryProcedureMatchesPosition(scope Range, position *Position) bool {
	if position == nil {
		return true
	}
	return comparePosition(scope.Start, *position) <= 0 && comparePosition(*position, scope.End) <= 0
}

func rebaseRange(source Range, oldStart, newStart Position) Range {
	return Range{Start: rebasePosition(source.Start, oldStart, newStart), End: rebasePosition(source.End, oldStart, newStart)}
}
