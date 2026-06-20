package intel

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"sort"
	"strings"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/lint"
	"github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vba/userforms"
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

type Completion struct {
	Label         string
	Kind          string
	Detail        string
	Documentation string
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
	issues, err := lint.Linter{RootDir: a.RootDir, Config: a.Config}.LintSource(doc.Path, []byte(doc.Source))
	if err != nil {
		return []Diagnostic{lineDiagnostic("VBA000", "error", 0, err.Error())}
	}
	out := make([]Diagnostic, 0, len(issues))
	for _, issue := range issues {
		out = append(out, Diagnostic{
			Code:     issue.Code,
			Severity: issue.Severity,
			Source:   "xlflow",
			Message:  issue.Message,
			Range:    issueRange(doc.Source, issue.Line, issue.Column),
		})
	}
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
	out := symbolsFromFile(file, doc.URI)
	out = append(out, a.formControlSymbols(doc)...)
	return out, nil
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
	openKeys := make(map[string]bool, len(open))
	for _, doc := range open {
		for _, key := range a.workspacePathKeys(doc.Path) {
			openKeys[key] = true
		}
	}
	var out []Symbol
	for _, file := range result.Files {
		if hasAnyPathKey(openKeys, a.workspacePathKeys(file.Path)) {
			continue
		}
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

func (a Analyzer) workspacePathKeys(path string) []string {
	keys := []string{}
	if key := pathKey(path); key != "" {
		keys = append(keys, key)
	}
	if strings.TrimSpace(path) == "" || strings.TrimSpace(a.RootDir) == "" {
		return keys
	}
	if filepath.IsAbs(path) {
		if rel, err := filepath.Rel(a.RootDir, path); err == nil {
			if key := pathKey(rel); key != "" {
				keys = append(keys, key)
			}
		}
		return keys
	}
	if key := pathKey(filepath.Join(a.RootDir, filepath.FromSlash(path))); key != "" {
		keys = append(keys, key)
	}
	return keys
}

func hasAnyPathKey(set map[string]bool, keys []string) bool {
	for _, key := range keys {
		if set[key] {
			return true
		}
	}
	return false
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

func (a Analyzer) References(doc Document, pos Position, open []Document, includeDeclaration bool, uriForPath func(string) string) ([]Location, error) {
	word, _ := WordAt(doc.Source, pos)
	if word == "" {
		return nil, nil
	}
	docs, err := a.workspaceDocuments(open)
	if err != nil {
		return nil, err
	}
	declarations := map[string]bool{}
	var declarationRanges []Location
	if !includeDeclaration {
		defs, err := a.Definition(doc, pos, open, nil)
		if err != nil {
			return nil, err
		}
		for _, def := range defs {
			declarations[locationKey(def.Path, def.Range)] = true
			declarationRanges = append(declarationRanges, def)
		}
		for _, candidate := range docs {
			syms, err := a.DocumentSymbols(candidate)
			if err != nil {
				continue
			}
			for _, sym := range syms {
				if !strings.EqualFold(sym.Name, word) {
					continue
				}
				declarations[locationKey(candidate.Path, sym.Selection)] = true
				declarations[locationKey(candidate.URI, sym.Selection)] = true
				declarationRanges = append(declarationRanges, Location{URI: candidate.URI, Path: candidate.Path, Range: sym.Range})
			}
		}
	}
	var out []Location
	for _, candidate := range docs {
		for _, r := range identifierRanges(candidate.Source, word) {
			if !includeDeclaration && (declarations[locationKey(candidate.Path, r)] || declarations[locationKey(candidate.URI, r)]) {
				continue
			}
			if !includeDeclaration && containedInDeclaration(candidate, r, declarationRanges) {
				continue
			}
			uri := candidate.URI
			if uri == "" && uriForPath != nil {
				uri = uriForPath(candidate.Path)
			}
			out = append(out, Location{URI: uri, Path: candidate.Path, Range: r})
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].Path != out[j].Path {
			return out[i].Path < out[j].Path
		}
		if out[i].Range.Start.Line != out[j].Range.Start.Line {
			return out[i].Range.Start.Line < out[j].Range.Start.Line
		}
		return out[i].Range.Start.Character < out[j].Range.Start.Character
	})
	return out, nil
}

func containedInDeclaration(doc Document, r Range, declarations []Location) bool {
	for _, declaration := range declarations {
		sameDocument := pathKey(declaration.Path) == pathKey(doc.Path)
		if !sameDocument && declaration.URI != "" && doc.URI != "" {
			sameDocument = declaration.URI == doc.URI
		}
		if sameDocument && declaration.Range.Start.Line == r.Start.Line && rangeContains(declaration.Range, r) {
			return true
		}
	}
	return false
}

func rangeContains(outer, inner Range) bool {
	if comparePosition(outer.Start, inner.Start) > 0 {
		return false
	}
	return comparePosition(outer.End, inner.End) >= 0
}

func comparePosition(a, b Position) int {
	if a.Line != b.Line {
		return a.Line - b.Line
	}
	return a.Character - b.Character
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
	if typ, ok := a.inferWordTypeAt(doc, word, byteOffsetForPosition(doc.Source, pos)); ok {
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

func (a Analyzer) Completions(doc Document, pos Position, open []Document) ([]Completion, error) {
	line := lineAt(doc.Source, pos.Line)
	prefix := utf16Prefix(line, pos.Character)
	memberPrefix, receiverExpr, memberMode := memberCompletionContext(prefix)
	if memberMode {
		typ, ok := a.resolveDocumentExpressionTypeAt(doc, receiverExpr, byteOffsetForPosition(doc.Source, pos))
		if !ok {
			return nil, nil
		}
		return a.memberCompletions(typ, memberPrefix), nil
	}
	word, _ := WordAt(doc.Source, pos)
	if strings.TrimSpace(word) == "" {
		word = currentIdentifierPrefix(prefix)
	}
	items := a.globalCompletions(word)
	for _, control := range a.formControls(doc) {
		if completionMatches(control.Name, word) {
			items = append(items, Completion{
				Label:  control.Name,
				Kind:   "variable",
				Detail: control.Type,
			})
		}
	}
	syms, err := a.WorkspaceSymbols(open, word)
	if err != nil {
		return nil, err
	}
	for _, sym := range syms {
		items = append(items, Completion{Label: sym.Name, Kind: completionKindForSymbol(sym.Kind), Detail: sym.Detail})
	}
	return uniqueCompletions(items), nil
}

func (a Analyzer) inferWordType(doc Document, word string) (string, bool) {
	return a.inferWordTypeAt(doc, word, -1)
}

func (a Analyzer) inferWordTypeAt(doc Document, word string, offset int) (string, bool) {
	if strings.EqualFold(word, "Me") && a.isFormDocument(doc) {
		return "MSForms.UserForm", true
	}
	if control, ok := a.resolveFormControl(doc, word); ok {
		return control.Type, true
	}
	if typ, ok := a.DB.ResolveGlobal(word); ok {
		return typ.Name, true
	}
	var declared string
	declRe := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(word) + `\b(?:\s*\([^)]*\))?\s+As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_.]*)`)
	if typ, ok := bestTypeMatch(doc.Source, declRe, offset, 1); ok {
		declared = typ
		if !isObjectFallbackType(declared) {
			return declared, true
		}
	}
	newRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*New\s+([A-Za-z_][A-Za-z0-9_.]*)`)
	if typ, ok := bestTypeMatch(doc.Source, newRe, offset, 1); ok {
		return typ, true
	}
	createRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*CreateObject\s*\(\s*"([^"]+)"\s*\)`)
	if progID, ok := bestTypeMatch(doc.Source, createRe, offset, 1); ok {
		if typ, ok := a.DB.ResolveProgID(progID); ok {
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
	return a.resolveExpressionType(Document{}, expr, false)
}

func (a Analyzer) resolveDocumentExpressionTypeAt(doc Document, expr string, offset int) (string, bool) {
	return a.resolveExpressionTypeAt(doc, expr, true, offset)
}

func (a Analyzer) resolveExpressionType(doc Document, expr string, useDocument bool) (string, bool) {
	return a.resolveExpressionTypeAt(doc, expr, useDocument, -1)
}

func (a Analyzer) resolveExpressionTypeAt(doc Document, expr string, useDocument bool, offset int) (string, bool) {
	parts := splitMemberExpression(expr)
	if len(parts) == 0 {
		return "", false
	}
	base := strings.TrimSpace(parts[0])
	if idx := strings.Index(base, "("); idx >= 0 {
		base = strings.TrimSpace(base[:idx])
	}
	var current string
	formMode := useDocument && strings.EqualFold(base, "Me") && a.isFormDocument(doc)
	if formMode {
		current = "MSForms.UserForm"
	} else if typ, ok := a.DB.ResolveGlobal(base); ok {
		current = typ.Name
	} else if typ, ok := a.DB.ResolveType(base); ok {
		current = typ.Name
	} else if useDocument {
		if typ, ok := a.inferWordTypeAt(doc, base, offset); ok {
			current = typ
		} else {
			return "", false
		}
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
		args := ""
		if idx := strings.Index(member, "("); idx >= 0 {
			args = member[idx:]
			member = strings.TrimSpace(member[:idx])
		}
		if member == "" {
			continue
		}
		if formMode {
			if control, ok := a.resolveFormControl(doc, member); ok {
				current = control.Type
				formMode = false
				continue
			}
			if strings.EqualFold(member, "Controls") {
				if controlName := firstStringArgument(args); controlName != "" {
					if control, ok := a.resolveFormControl(doc, controlName); ok {
						current = control.Type
						formMode = false
						continue
					}
				}
			}
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

func bestTypeMatch(source string, re *regexp.Regexp, offset int, group int) (string, bool) {
	matches := re.FindAllStringSubmatchIndex(source, -1)
	bestStart := -1
	bestType := ""
	for _, match := range matches {
		if len(match) <= group*2+1 || match[group*2] < 0 || match[group*2+1] < 0 {
			continue
		}
		start := match[0]
		if offset >= 0 && start > offset {
			continue
		}
		if bestStart < 0 || start > bestStart {
			bestStart = start
			bestType = source[match[group*2]:match[group*2+1]]
		}
	}
	if bestType == "" && offset >= 0 {
		return bestTypeMatch(source, re, -1, group)
	}
	return bestType, bestType != ""
}

func (a Analyzer) collectionDefaultType(name string) (string, bool) {
	typ, ok := a.DB.ResolveType(name)
	if !ok || typ.ElementType == "" {
		return "", false
	}
	return typ.ElementType, true
}

func (a Analyzer) formControlSymbols(doc Document) []Symbol {
	controls := a.formControls(doc)
	out := make([]Symbol, 0, len(controls))
	for _, control := range controls {
		out = append(out, Symbol{
			Name:   control.Name,
			Kind:   "field",
			Detail: control.Name + " As " + control.Type,
			File:   firstNonEmpty(doc.URI, doc.Path),
			Module: a.formName(doc),
			Range: Range{
				Start: Position{Line: max(0, control.StartLine-1), Character: max(0, control.StartColumn-1)},
				End:   Position{Line: max(0, control.EndLine-1), Character: max(0, control.EndColumn-1)},
			},
			Selection: Range{
				Start: Position{Line: max(0, control.StartLine-1), Character: max(0, control.StartColumn-1)},
				End:   Position{Line: max(0, control.StartLine-1), Character: max(0, control.StartColumn-1+utf16Len(control.Name))},
			},
		})
	}
	return out
}

func (a Analyzer) formControls(doc Document) []userforms.Control {
	form := userforms.Parse(a.formSource(doc))
	return form.Controls
}

func (a Analyzer) resolveFormControl(doc Document, name string) (userforms.Control, bool) {
	for _, control := range a.formControls(doc) {
		if strings.EqualFold(control.Name, name) {
			return control, true
		}
	}
	return userforms.Control{}, false
}

func (a Analyzer) isFormDocument(doc Document) bool {
	return strings.EqualFold(doc.ModuleKind, "form") || strings.EqualFold(filepath.Ext(doc.Path), ".frm") || a.formSource(doc) != ""
}

func (a Analyzer) formName(doc Document) string {
	form := userforms.Parse(a.formSource(doc))
	if form.Name != "" {
		return form.Name
	}
	return strings.TrimSuffix(filepath.Base(doc.Path), filepath.Ext(doc.Path))
}

func (a Analyzer) formSource(doc Document) string {
	if strings.EqualFold(filepath.Ext(doc.Path), ".frm") {
		return doc.Source
	}
	for _, path := range a.candidateFormPaths(doc) {
		body, err := os.ReadFile(path)
		if err == nil {
			return string(body)
		}
	}
	return ""
}

func (a Analyzer) candidateFormPaths(doc Document) []string {
	name := strings.TrimSuffix(filepath.Base(doc.Path), filepath.Ext(doc.Path))
	var paths []string
	if a.RootDir != "" && a.Config.Src.Forms != "" {
		formsRoot := filepath.Join(a.RootDir, filepath.FromSlash(a.Config.Src.Forms))
		paths = append(paths, filepath.Join(formsRoot, name+".frm"))
		if parent := filepath.Base(filepath.Dir(doc.Path)); strings.EqualFold(parent, "code") {
			paths = append(paths, filepath.Join(filepath.Dir(filepath.Dir(doc.Path)), name+".frm"))
		}
	}
	return paths
}

func (a Analyzer) workspaceDocuments(open []Document) ([]Document, error) {
	out := make([]Document, 0, len(open))
	seen := map[string]bool{}
	for _, doc := range open {
		key := pathKey(doc.Path)
		if key == "" {
			key = doc.URI
		}
		seen[key] = true
		out = append(out, doc)
	}
	dirs := []struct {
		path string
		kind string
	}{
		{a.Config.Src.Modules, "standard"},
		{a.Config.Src.Classes, "class"},
		{a.Config.Src.Forms, "form"},
		{a.Config.Src.Workbook, "document"},
	}
	for _, dir := range dirs {
		if strings.TrimSpace(dir.path) == "" {
			continue
		}
		root := filepath.Join(a.RootDir, filepath.FromSlash(dir.path))
		err := filepath.WalkDir(root, func(path string, d os.DirEntry, walkErr error) error {
			if walkErr != nil {
				if os.IsNotExist(walkErr) {
					return nil
				}
				return walkErr
			}
			if d.IsDir() {
				return nil
			}
			ext := strings.ToLower(filepath.Ext(path))
			if ext != ".bas" && ext != ".cls" && ext != ".frm" {
				return nil
			}
			key := pathKey(path)
			if seen[key] {
				return nil
			}
			body, err := os.ReadFile(path)
			if err != nil {
				return err
			}
			seen[key] = true
			out = append(out, Document{Path: path, Source: string(body), ModuleKind: dir.kind})
			return nil
		})
		if err != nil && !os.IsNotExist(err) {
			return nil, err
		}
	}
	return out, nil
}

func identifierRanges(source, name string) []Range {
	if strings.TrimSpace(name) == "" {
		return nil
	}
	var out []Range
	for lineNo, line := range normalizedLines(source) {
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
			if strings.EqualFold(line[start:end], name) {
				out = append(out, Range{
					Start: Position{Line: lineNo, Character: utf16Len(line[:start])},
					End:   Position{Line: lineNo, Character: utf16Len(line[:end])},
				})
			}
			start = end
		}
	}
	return out
}

func codeLimit(line string) int {
	inString := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '\'':
			if !inString {
				return i
			}
		}
	}
	return len(line)
}

func (a Analyzer) memberCompletions(receiverType, prefix string) []Completion {
	var out []Completion
	for _, member := range a.DB.Members(receiverType) {
		if !completionMatches(member.Name, prefix) {
			continue
		}
		kind := "method"
		if member.ReadOnly || member.WriteOnly || member.ReturnType != "" {
			kind = "property"
		}
		out = append(out, Completion{
			Label:         member.Name,
			Kind:          kind,
			Detail:        memberDetail(member),
			Documentation: member.Summary,
		})
	}
	return uniqueCompletions(out)
}

func (a Analyzer) globalCompletions(prefix string) []Completion {
	var out []Completion
	for _, typ := range a.DB.TypeNames() {
		if completionMatches(typ, prefix) {
			out = append(out, Completion{Label: typ, Kind: "type", Detail: "type"})
		}
	}
	for _, constant := range a.DB.ConstantsList() {
		if completionMatches(constant.Name, prefix) {
			out = append(out, Completion{
				Label:         constant.Name,
				Kind:          "constant",
				Detail:        firstNonEmpty(constant.EnumGroup, constant.Type, constant.Kind),
				Documentation: constant.Summary,
			})
		}
	}
	for _, global := range a.DB.GlobalsList() {
		if completionMatches(global.Name, prefix) {
			out = append(out, Completion{
				Label:  global.Name,
				Kind:   "variable",
				Detail: global.ReturnType,
			})
		}
	}
	return uniqueCompletions(out)
}

func memberCompletionContext(prefix string) (memberPrefix string, receiverExpr string, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !strings.HasSuffix(beforeWord, ".") {
		return "", "", false
	}
	receiver := expressionBefore(strings.TrimSuffix(beforeWord, "."))
	if receiver == "" {
		return "", "", false
	}
	return wordPrefix, receiver, true
}

func currentIdentifierPrefix(prefix string) string {
	end := len(prefix)
	start := end
	for start > 0 {
		r, size := lastRune(prefix[:start])
		if !isIdentRune(r) {
			break
		}
		start -= size
	}
	return prefix[start:end]
}

func completionMatches(label, prefix string) bool {
	prefix = strings.TrimSpace(prefix)
	return prefix == "" || strings.HasPrefix(strings.ToLower(label), strings.ToLower(prefix))
}

func uniqueCompletions(items []Completion) []Completion {
	seen := map[string]bool{}
	out := items[:0]
	for _, item := range items {
		if item.Label == "" {
			continue
		}
		key := strings.ToLower(item.Kind + "\x00" + item.Label)
		if seen[key] {
			continue
		}
		seen[key] = true
		out = append(out, item)
	}
	sort.SliceStable(out, func(i, j int) bool {
		if out[i].Label != out[j].Label {
			return strings.ToLower(out[i].Label) < strings.ToLower(out[j].Label)
		}
		return out[i].Kind < out[j].Kind
	})
	return out
}

func memberDetail(member vbadb.MemberInfo) string {
	if member.ReturnType == "" {
		return member.Name
	}
	return member.Name + " As " + member.ReturnType
}

func completionKindForSymbol(kind string) string {
	switch strings.ToLower(kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return "function"
	case "const":
		return "constant"
	case "module_variable", "local_variable", "field":
		return "variable"
	case "class", "module", "enum":
		return "type"
	default:
		return "symbol"
	}
}

func locationKey(path string, r Range) string {
	return fmt.Sprintf("%s:%d:%d:%d:%d", pathKey(path), r.Start.Line, r.Start.Character, r.End.Line, r.End.Character)
}

func pathKey(path string) string {
	if strings.TrimSpace(path) == "" {
		return ""
	}
	clean := filepath.Clean(path)
	if runtime.GOOS == "windows" {
		clean = strings.ToLower(clean)
	}
	return clean
}

func symbolsFromFile(file symbols.FileResult, uri string) []Symbol {
	out := make([]Symbol, 0, len(file.Symbols))
	for _, sym := range file.Symbols {
		if strings.TrimSpace(sym.Name) == "" {
			continue
		}
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

func issueRange(source string, oneLine int, oneColumn int) Range {
	line := max(0, oneLine-1)
	lines := normalizedLines(source)
	if len(lines) == 0 {
		return Range{Start: Position{Line: line, Character: 0}, End: Position{Line: line, Character: 1}}
	}
	if line >= len(lines) {
		line = len(lines) - 1
	}
	start := 0
	if oneColumn > 0 {
		start = oneColumn - 1
	}
	text := lines[line]
	if start > len(text) {
		start = len(text)
	}
	for start > 0 && start < len(text) && !utf8.RuneStart(text[start]) {
		start--
	}
	character := utf16Len(text[:start])
	return Range{
		Start: Position{Line: line, Character: character},
		End:   Position{Line: line, Character: character + 1},
	}
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

func byteOffsetForPosition(source string, pos Position) int {
	lines := normalizedLines(source)
	if pos.Line < 0 {
		return 0
	}
	offset := 0
	for lineNo, line := range lines {
		if lineNo == pos.Line {
			idx := byteIndexForUTF16(line, pos.Character)
			if idx > len(line) {
				idx = len(line)
			}
			return offset + idx
		}
		offset += len(line) + 1
	}
	return len(source)
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

func firstStringArgument(args string) string {
	start := strings.Index(args, `"`)
	if start < 0 {
		return ""
	}
	var b strings.Builder
	for i := start + 1; i < len(args); i++ {
		if args[i] != '"' {
			b.WriteByte(args[i])
			continue
		}
		if i+1 < len(args) && args[i+1] == '"' {
			b.WriteByte('"')
			i++
			continue
		}
		return b.String()
	}
	return ""
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
	if c.EnumGroup != "" {
		b.WriteString("\n")
		b.WriteString(c.EnumGroup)
	}
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

func isIdentStartRune(r rune) bool {
	return r == '_' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z'
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
