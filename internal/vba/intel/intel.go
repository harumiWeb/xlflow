package intel

import (
	"fmt"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

type Analyzer struct {
	RootDir string
	Config  config.Config
	DB      *vbadb.DB
}

type Document struct {
	URI        string
	Path       string
	Source     string
	ModuleKind string
}

type Position struct {
	Line      int
	Character int
}

type Range struct {
	Start Position
	End   Position
}

type Diagnostic struct {
	Code     string
	Severity string
	Source   string
	Message  string
	Range    Range
}

type Symbol struct {
	Name      string
	Kind      string
	Detail    string
	File      string
	Module    string
	Range     Range
	Selection Range
}

type Location struct {
	URI   string
	Path  string
	Range Range
}

type Hover struct {
	Contents string
	Range    Range
}

func (a Analyzer) Check() error {
	if a.DB == nil {
		return fmt.Errorf("VBA type database is not loaded")
	}
	if _, ok := a.DB.ResolveType("Excel.Application"); !ok {
		return fmt.Errorf("built-in type database is missing Excel.Application")
	}
	parser, err := ast.NewParser()
	if err != nil {
		return err
	}
	parser.Close()
	return nil
}

func (a Analyzer) Diagnostics(doc Document) []Diagnostic {
	parser, err := ast.NewParser()
	if err != nil {
		return []Diagnostic{lineDiagnostic("VBA000", "error", 0, err.Error())}
	}
	defer parser.Close()
	parsed := parser.Parse(doc.Path, []byte(doc.Source))
	defer parsed.Close()
	var out []Diagnostic
	if parsed.HasError || parsed.HasMissing {
		out = append(out, Diagnostic{
			Code:     "VB014",
			Severity: "error",
			Source:   "xlflow",
			Message:  "VBA parser recovered from syntax errors; inspect this source before pushing to Excel.",
			Range:    nodeRange(parsed.Root, doc.Source),
		})
	}
	if !hasOptionExplicit(doc.Source) {
		out = append(out, lineDiagnostic("VB001", "error", 0, "Missing Option Explicit."))
	}
	out = append(out, implicitVariantDiagnostics(doc.Source)...)
	return out
}

func (a Analyzer) DocumentSymbols(doc Document) ([]Symbol, error) {
	file, err := symbols.InspectSource(symbols.SourceOptions{
		RootDir:        a.RootDir,
		Path:           doc.Path,
		ModuleKind:     doc.ModuleKind,
		IncludePrivate: true,
		IncludeLabels:  true,
	}, []byte(doc.Source))
	if err != nil {
		return nil, err
	}
	return symbolsFromFile(file, doc.URI), nil
}

func (a Analyzer) WorkspaceSymbols(open []Document, query string) ([]Symbol, error) {
	result, err := symbols.Inspect(symbols.Options{
		RootDir:        a.RootDir,
		Config:         a.Config,
		IncludePrivate: true,
		IncludeLabels:  false,
	})
	if err != nil {
		return nil, err
	}
	var out []Symbol
	for _, file := range result.Files {
		out = append(out, symbolsFromFile(file, "")...)
	}
	for _, doc := range open {
		docSyms, err := a.DocumentSymbols(doc)
		if err != nil {
			continue
		}
		out = append(out, docSyms...)
	}
	query = strings.ToLower(strings.TrimSpace(query))
	if query != "" {
		filtered := out[:0]
		for _, sym := range out {
			if strings.Contains(strings.ToLower(sym.Name), query) || strings.Contains(strings.ToLower(sym.Module+"."+sym.Name), query) {
				filtered = append(filtered, sym)
			}
		}
		out = filtered
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].File != out[j].File {
			return out[i].File < out[j].File
		}
		if out[i].Range.Start.Line != out[j].Range.Start.Line {
			return out[i].Range.Start.Line < out[j].Range.Start.Line
		}
		return out[i].Name < out[j].Name
	})
	return out, nil
}

func (a Analyzer) Definition(doc Document, pos Position, open []Document, uriForPath func(string) string) ([]Location, error) {
	word, _ := WordAt(doc.Source, pos)
	if word == "" {
		return nil, nil
	}
	syms, err := a.WorkspaceSymbols(open, word)
	if err != nil {
		return nil, err
	}
	var out []Location
	for _, sym := range syms {
		if !strings.EqualFold(sym.Name, word) {
			continue
		}
		uri := sym.File
		if uriForPath != nil {
			uri = uriForPath(sym.File)
		}
		out = append(out, Location{URI: uri, Path: sym.File, Range: sym.Selection})
	}
	return out, nil
}

func (a Analyzer) Hover(doc Document, pos Position, open []Document) (*Hover, error) {
	word, r := WordAt(doc.Source, pos)
	if word == "" {
		return nil, nil
	}
	if typ, ok := a.DB.ResolveType(word); ok {
		return &Hover{Contents: typeHover(typ), Range: r}, nil
	}
	if constant, ok := a.DB.ResolveConstant(word); ok {
		return &Hover{Contents: constantHover(constant), Range: r}, nil
	}
	if typ, ok := a.inferWordType(doc, word); ok {
		if dbType, found := a.DB.ResolveType(typ); found {
			return &Hover{Contents: typeHover(dbType), Range: r}, nil
		}
		return &Hover{Contents: "```vb\n" + word + " As " + typ + "\n```", Range: r}, nil
	}
	syms, err := a.WorkspaceSymbols(open, word)
	if err != nil {
		return nil, err
	}
	for _, sym := range syms {
		if strings.EqualFold(sym.Name, word) {
			detail := sym.Detail
			if detail == "" {
				detail = sym.Kind + " " + sym.Name
			}
			return &Hover{Contents: "```vb\n" + detail + "\n```", Range: r}, nil
		}
	}
	if typ, ok := a.inferExpressionType(doc.Source, pos); ok {
		if dbType, found := a.DB.ResolveType(typ); found {
			return &Hover{Contents: typeHover(dbType), Range: r}, nil
		}
	}
	return nil, nil
}

func (a Analyzer) inferWordType(doc Document, word string) (string, bool) {
	if typ, ok := a.DB.ResolveGlobal(word); ok {
		return typ.Name, true
	}
	var declared string
	declRe := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(word) + `\b(?:\s*\([^)]*\))?\s+As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_.]*)`)
	if m := declRe.FindStringSubmatch(doc.Source); len(m) > 1 {
		declared = m[1]
		if !isObjectFallbackType(declared) {
			return declared, true
		}
	}
	newRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*New\s+([A-Za-z_][A-Za-z0-9_.]*)`)
	if m := newRe.FindStringSubmatch(doc.Source); len(m) > 1 {
		return m[1], true
	}
	createRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*CreateObject\s*\(\s*"([^"]+)"\s*\)`)
	if m := createRe.FindStringSubmatch(doc.Source); len(m) > 1 {
		if typ, ok := a.DB.ResolveProgID(m[1]); ok {
			return typ.Name, true
		}
	}
	if declared != "" {
		return declared, true
	}
	return "", false
}

func isObjectFallbackType(name string) bool {
	return strings.EqualFold(name, "Object") || strings.EqualFold(name, "Variant")
}

func (a Analyzer) inferExpressionType(source string, pos Position) (string, bool) {
	line := lineAt(source, pos.Line)
	if line == "" {
		return "", false
	}
	prefix := utf16Prefix(line, pos.Character)
	expr := expressionBefore(prefix)
	if expr == "" {
		return "", false
	}
	return a.ResolveExpressionType(expr)
}

func (a Analyzer) ResolveExpressionType(expr string) (string, bool) {
	parts := splitMemberExpression(expr)
	if len(parts) == 0 {
		return "", false
	}
	base := strings.TrimSpace(parts[0])
	if idx := strings.Index(base, "("); idx >= 0 {
		base = strings.TrimSpace(base[:idx])
	}
	var current string
	if typ, ok := a.DB.ResolveGlobal(base); ok {
		current = typ.Name
	} else if typ, ok := a.DB.ResolveType(base); ok {
		current = typ.Name
	} else {
		return "", false
	}
	if strings.Contains(parts[0], "(") {
		if typ, ok := a.collectionDefaultType(current); ok {
			current = typ
		}
	}
	for _, raw := range parts[1:] {
		member := strings.TrimSpace(raw)
		called := strings.Contains(member, "(")
		if idx := strings.Index(member, "("); idx >= 0 {
			member = strings.TrimSpace(member[:idx])
		}
		if member == "" {
			continue
		}
		if info, ok := a.DB.ResolveMember(current, member); ok && info.ReturnType != "" {
			current = info.ReturnType
		} else {
			return "", false
		}
		if called {
			if typ, ok := a.collectionDefaultType(current); ok {
				current = typ
			}
		}
	}
	return current, true
}

func (a Analyzer) collectionDefaultType(name string) (string, bool) {
	typ, ok := a.DB.ResolveType(name)
	if !ok || typ.ElementType == "" {
		return "", false
	}
	return typ.ElementType, true
}

func symbolsFromFile(file symbols.FileResult, uri string) []Symbol {
	out := make([]Symbol, 0, len(file.Symbols))
	for _, sym := range file.Symbols {
		out = append(out, Symbol{
			Name:   sym.Name,
			Kind:   sym.Kind,
			Detail: firstNonEmpty(sym.Signature, sym.Kind+" "+sym.Name),
			File:   firstNonEmpty(uri, file.Path, sym.File),
			Module: sym.Module,
			Range: Range{
				Start: Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1)},
				End:   Position{Line: sym.EndLine - 1, Character: max(0, sym.EndColumn-1)},
			},
			Selection: Range{
				Start: Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1)},
				End:   Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1+len([]rune(sym.Name)))},
			},
		})
	}
	return out
}

func lineDiagnostic(code, severity string, zeroLine int, msg string) Diagnostic {
	return Diagnostic{
		Code:     code,
		Severity: severity,
		Source:   "xlflow",
		Message:  msg,
		Range: Range{
			Start: Position{Line: zeroLine, Character: 0},
			End:   Position{Line: zeroLine, Character: 1},
		},
	}
}

func implicitVariantDiagnostics(source string) []Diagnostic {
	re := regexp.MustCompile(`(?i)^\s*(?:Dim|Private|Public|Static)\s+(.+)$`)
	var out []Diagnostic
	for i, line := range normalizedLines(source) {
		m := re.FindStringSubmatch(line)
		if len(m) == 0 {
			continue
		}
		parts := strings.Split(m[1], ",")
		offset := strings.Index(line, m[1])
		for _, part := range parts {
			part = strings.TrimSpace(part)
			if part == "" || strings.Contains(strings.ToLower(part), " as ") {
				offset += len(part) + 1
				continue
			}
			name := strings.FieldsFunc(part, func(r rune) bool {
				return !isIdentRune(r)
			})
			if len(name) == 0 {
				continue
			}
			col := strings.Index(line[offset:], name[0])
			if col < 0 {
				col = 0
			}
			start := offset + col
			out = append(out, Diagnostic{
				Code:     "VB005",
				Severity: "warning",
				Source:   "xlflow",
				Message:  "Declare an explicit type with As <Type>.",
				Range: Range{
					Start: Position{Line: i, Character: utf16Len(line[:start])},
					End:   Position{Line: i, Character: utf16Len(line[:start+len(name[0])])},
				},
			})
			offset += len(part) + 1
		}
	}
	return out
}

func hasOptionExplicit(source string) bool {
	for _, line := range normalizedLines(source) {
		clean := strings.TrimSpace(strings.Split(line, "'")[0])
		if strings.EqualFold(clean, "Option Explicit") {
			return true
		}
		if clean != "" && !strings.HasPrefix(strings.ToLower(clean), "attribute ") {
			return false
		}
	}
	return false
}

func nodeRange(node interface {
	StartByte() uint
	EndByte() uint
}, source string) Range {
	if node == nil {
		return Range{Start: Position{}, End: Position{Character: 1}}
	}
	return byteRangeToRange(source, int(node.StartByte()), int(node.EndByte()))
}

func byteRangeToRange(source string, start, end int) Range {
	if start < 0 {
		start = 0
	}
	if end < start {
		end = start
	}
	if end > len(source) {
		end = len(source)
	}
	return Range{Start: byteOffsetToPosition(source, start), End: byteOffsetToPosition(source, end)}
}

func byteOffsetToPosition(source string, offset int) Position {
	if offset < 0 {
		offset = 0
	}
	if offset > len(source) {
		offset = len(source)
	}
	line := 0
	lineStart := 0
	for i := 0; i < offset; i++ {
		if source[i] == '\n' {
			line++
			lineStart = i + 1
		}
	}
	return Position{Line: line, Character: utf16Len(source[lineStart:offset])}
}

func WordAt(source string, pos Position) (string, Range) {
	line := lineAt(source, pos.Line)
	if line == "" {
		return "", Range{Start: pos, End: pos}
	}
	byteCol := byteIndexForUTF16(line, pos.Character)
	if byteCol > len(line) {
		byteCol = len(line)
	}
	start := byteCol
	for start > 0 {
		r, size := lastRune(line[:start])
		if !isIdentRune(r) {
			break
		}
		start -= size
	}
	end := byteCol
	for end < len(line) {
		r, size := firstRune(line[end:])
		if !isIdentRune(r) {
			break
		}
		end += size
	}
	word := line[start:end]
	return word, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(line[:start])},
		End:   Position{Line: pos.Line, Character: utf16Len(line[:end])},
	}
}

func normalizedLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func lineAt(source string, zeroLine int) string {
	lines := normalizedLines(source)
	if zeroLine < 0 || zeroLine >= len(lines) {
		return ""
	}
	return lines[zeroLine]
}

func utf16Prefix(line string, character int) string {
	idx := byteIndexForUTF16(line, character)
	if idx > len(line) {
		idx = len(line)
	}
	return line[:idx]
}

func byteIndexForUTF16(s string, character int) int {
	if character <= 0 {
		return 0
	}
	units := 0
	for idx, r := range s {
		next := units + len(utf16.Encode([]rune{r}))
		if next > character {
			return idx
		}
		units = next
	}
	return len(s)
}

func utf16Len(s string) int {
	return len(utf16.Encode([]rune(s)))
}

func expressionBefore(prefix string) string {
	prefix = strings.TrimRight(prefix, " \t")
	if prefix == "" {
		return ""
	}
	start := len(prefix)
	depth := 0
	for start > 0 {
		r, size := lastRune(prefix[:start])
		if r == ')' {
			depth++
		} else if r == '(' && depth > 0 {
			depth--
		} else if depth == 0 && !isExprRune(r) {
			break
		}
		start -= size
	}
	return strings.TrimSpace(prefix[start:])
}

func splitMemberExpression(expr string) []string {
	var parts []string
	start := 0
	depth := 0
	inString := false
	for i := 0; i < len(expr); i++ {
		switch expr[i] {
		case '"':
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		case '.':
			if !inString && depth == 0 {
				parts = append(parts, strings.TrimSpace(expr[start:i]))
				start = i + 1
			}
		}
	}
	if tail := strings.TrimSpace(expr[start:]); tail != "" {
		parts = append(parts, tail)
	}
	return parts
}

func typeHover(typ vbadb.TypeInfo) string {
	var b strings.Builder
	b.WriteString("```vb\n")
	b.WriteString(typ.Name)
	if typ.Kind != "" {
		b.WriteString(" As ")
		b.WriteString(typ.Kind)
	}
	b.WriteString("\n```")
	if typ.Summary != "" {
		b.WriteString("\n")
		b.WriteString(typ.Summary)
	}
	if typ.ElementType != "" {
		b.WriteString("\n\nCollection element: `")
		b.WriteString(typ.ElementType)
		b.WriteString("`")
	}
	return b.String()
}

func constantHover(c vbadb.ConstantInfo) string {
	var b strings.Builder
	b.WriteString("```vb\nConst ")
	b.WriteString(c.Name)
	if c.Type != "" {
		b.WriteString(" As ")
		b.WriteString(c.Type)
	}
	if c.Value != "" {
		b.WriteString(" = ")
		b.WriteString(c.Value)
	}
	b.WriteString("\n```")
	if c.Summary != "" {
		b.WriteString("\n")
		b.WriteString(c.Summary)
	}
	return b.String()
}

func isIdentRune(r rune) bool {
	return r == '_' || r == '$' || r == '%' || r == '&' || r == '!' || r == '#' || r == '@' || r == '^' ||
		r >= '0' && r <= '9' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z'
}

func isExprRune(r rune) bool {
	return isIdentRune(r) || r == '.' || r == '(' || r == ')' || r == '"' || r == ',' || r == ' '
}

func firstRune(s string) (rune, int) {
	return utf8.DecodeRuneInString(s)
}

func lastRune(s string) (rune, int) {
	return utf8.DecodeLastRuneInString(s)
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func DisplayPath(root, path string) string {
	if root != "" {
		if rel, err := filepath.Rel(root, path); err == nil {
			return filepath.ToSlash(rel)
		}
	}
	return filepath.ToSlash(path)
}
