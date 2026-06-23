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
	RootDir              string
	Config               config.Config
	DB                   *vbadb.DB
	WorkspaceSymbolsFunc func(open []Document, query string) ([]Symbol, error)
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
	Name       string
	Kind       string
	Detail     string
	ReturnType string
	Parameters []Parameter
	File       string
	Module     string
	ModuleKind string
	Parent     string
	Visibility string
	Range      Range
	Selection  Range
}

type Parameter struct {
	Name     string
	Type     string
	Optional bool
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
	InsertText    string
	Snippet       bool
	ReplaceRange  *Range
}

type SignatureHelp struct {
	Signatures      []Signature
	ActiveSignature int
	ActiveParameter int
}

type Signature struct {
	Label         string
	Parameters    []Parameter
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
	out = append(out, a.argumentDiagnostics(doc)...)
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
	if a.WorkspaceSymbolsFunc != nil {
		return a.WorkspaceSymbolsFunc(open, query)
	}
	return a.workspaceSymbols(open, query)
}

func (a Analyzer) workspaceSymbols(open []Document, query string) ([]Symbol, error) {
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
	syms, err := a.definitionSymbols(doc, pos, open, word)
	if err != nil {
		return nil, err
	}
	var out []Location
	for _, sym := range syms {
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
	defSyms, err := a.definitionSymbols(doc, pos, open, word)
	if err != nil {
		return nil, err
	}
	var localScope *Range
	if hasLocalDefinitionSymbol(defSyms) {
		if scope, ok := currentProcedureRangeForDocument(doc, pos); ok {
			localScope = &scope
			docs = []Document{doc}
		}
	}
	declarations := map[string]bool{}
	var declarationRanges []Location
	if !includeDeclaration {
		for _, sym := range defSyms {
			declarations[locationKey(sym.File, sym.Selection)] = true
			declarationRanges = append(declarationRanges, Location{Path: sym.File, Range: sym.Range})
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
			if localScope != nil && !rangeContains(*localScope, r) {
				continue
			}
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

func (a Analyzer) definitionSymbols(doc Document, pos Position, open []Document, word string) ([]Symbol, error) {
	syms, err := a.WorkspaceSymbols(open, word)
	if err != nil {
		return nil, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	var out []Symbol
	for _, sym := range syms {
		if !strings.EqualFold(sym.Name, word) || !a.visibleDefinitionSymbol(doc, currentProcedure, sym) {
			continue
		}
		out = append(out, sym)
	}
	if local := localDefinitionSymbols(out); len(local) > 0 {
		return local, nil
	}
	return out, nil
}

func (a Analyzer) visibleDefinitionSymbol(doc Document, currentProcedure string, sym Symbol) bool {
	if sym.Parent != "" {
		return currentProcedure != "" && strings.EqualFold(sym.Parent, currentProcedure) && a.sameDocumentSymbol(doc, sym)
	}
	if a.sameDocumentSymbol(doc, sym) {
		return true
	}
	return !strings.EqualFold(sym.Visibility, "Private")
}

func localDefinitionSymbols(syms []Symbol) []Symbol {
	var out []Symbol
	for _, sym := range syms {
		if isLocalSymbol(sym) {
			out = append(out, sym)
		}
	}
	return out
}

func hasLocalDefinitionSymbol(syms []Symbol) bool {
	return len(localDefinitionSymbols(syms)) > 0
}

func isLocalSymbol(sym Symbol) bool {
	return strings.EqualFold(sym.Kind, "local_variable") || strings.EqualFold(sym.Kind, "parameter") || (sym.Parent != "" && strings.EqualFold(sym.Kind, "const"))
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
	if hover, ok := a.memberHover(doc, word, r, byteOffsetForPosition(doc.Source, pos)); ok {
		return hover, nil
	}
	if control, ok := a.resolveFormControl(doc, word); ok {
		return &Hover{Contents: variableHover(word, control.Type, "UserForm control"), Range: r}, nil
	}
	if typ, ok := a.DB.ResolveType(word); ok {
		return &Hover{Contents: typeHover(typ, "built-in type database"), Range: r}, nil
	}
	if constant, ok := a.DB.ResolveConstant(word); ok {
		return &Hover{Contents: constantHover(constant), Range: r}, nil
	}
	if inferred, ok := a.inferWordTypeInfoAt(doc, word, byteOffsetForPosition(doc.Source, pos)); ok {
		typ := inferred.Type
		if dbType, found := a.DB.ResolveType(typ); found {
			return &Hover{Contents: variableHover(word, dbType.Name, inferred.Source), Range: r}, nil
		}
		return &Hover{Contents: variableHover(word, typ, inferred.Source), Range: r}, nil
	}
	syms, err := a.definitionSymbols(doc, pos, open, word)
	if err != nil {
		return nil, err
	}
	if len(syms) > 0 {
		sym := syms[0]
		detail := sym.Detail
		if detail == "" {
			detail = sym.Kind + " " + sym.Name
		}
		return &Hover{Contents: symbolHover(detail, symbolSource(sym)), Range: r}, nil
	}
	if typ, ok := a.inferExpressionType(doc.Source, pos); ok {
		if dbType, found := a.DB.ResolveType(typ); found {
			return &Hover{Contents: typeHover(dbType, "inferred expression"), Range: r}, nil
		}
	}
	return nil, nil
}

func (a Analyzer) Completions(doc Document, pos Position, open []Document) ([]Completion, error) {
	line := lineAt(doc.Source, pos.Line)
	prefix := utf16Prefix(line, pos.Character)
	if progIDPrefix, replaceRange, ok := createObjectProgIDCompletionContext(prefix, pos); ok {
		return a.progIDCompletions(progIDPrefix, replaceRange), nil
	}
	if insideOpenString(prefix) {
		return nil, nil
	}
	if memberPrefix, receiverType, ok := a.withBlockMemberCompletionContext(doc, pos, prefix); ok {
		return a.memberCompletions(receiverType, memberPrefix), nil
	}
	memberPrefix, receiverExpr, memberMode := memberCompletionContext(prefix)
	if memberMode {
		typ, ok := a.resolveDocumentExpressionTypeAt(doc, receiverExpr, byteOffsetForPosition(doc.Source, pos))
		if ok {
			if strings.EqualFold(typ, "Object") && sheetsDefaultExpression(receiverExpr) {
				typ = "Excel.Worksheet"
			}
			return a.memberCompletions(typ, memberPrefix), nil
		}
		return a.moduleMemberCompletions(open, receiverExpr, memberPrefix)
	}
	if typePrefix, replaceRange, ok := typeCompletionContext(prefix, pos); ok {
		return a.typeCompletions(typePrefix, replaceRange, doc, open)
	}
	if callPrefix, replaceRange, ok := callCompletionContext(prefix, pos); ok {
		return a.callCompletions(callPrefix, replaceRange, doc, pos, open)
	}
	if setPrefix, replaceRange, ok := setRHSCompletionContext(prefix, pos); ok {
		return a.setRHSCompletions(setPrefix, replaceRange, doc, pos, open)
	}
	if valuePrefix, replaceRange, ok := valueRHSCompletionContext(prefix, pos); ok {
		return a.valueRHSCompletions(valuePrefix, replaceRange, doc, pos, open)
	}
	if conditionPrefix, replaceRange, ok := conditionCompletionContext(prefix, pos); ok {
		return a.valueRHSCompletions(conditionPrefix, replaceRange, doc, pos, open)
	}
	if eachPrefix, replaceRange, ok := forEachInCompletionContext(prefix, pos); ok {
		return a.valueRHSCompletions(eachPrefix, replaceRange, doc, pos, open)
	}
	word, _ := WordAt(doc.Source, pos)
	if strings.TrimSpace(word) == "" {
		word = currentIdentifierPrefix(prefix)
	}
	items := a.syntaxCompletions(doc, pos, prefix)
	items = append(items, a.globalCompletions(word)...)
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
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	for _, sym := range syms {
		if !a.visibleCompletionSymbol(doc, currentProcedure, sym) {
			continue
		}
		items = append(items, Completion{Label: sym.Name, Kind: completionKindForSymbol(sym.Kind), Detail: sym.Detail})
	}
	return uniqueCompletions(items), nil
}

func (a Analyzer) SignatureHelp(doc Document, pos Position, open []Document) (*SignatureHelp, error) {
	line := lineAt(doc.Source, pos.Line)
	prefix := utf16Prefix(line, pos.Character)
	call, ok := callContextFromPrefix(prefix)
	if !ok {
		return nil, nil
	}
	sig, ok, err := a.resolveCallSignature(doc, call.Target, pos, open)
	if err != nil || !ok {
		return nil, err
	}
	active := call.ActiveParameter(sig.Parameters)
	return &SignatureHelp{
		Signatures:      []Signature{sig},
		ActiveSignature: 0,
		ActiveParameter: active,
	}, nil
}

type callContext struct {
	Target     string
	Arguments  []argument
	ActiveName string
	ActivePos  int
}

type argument struct {
	Name string
	Text string
}

func callContextFromPrefix(prefix string) (callContext, bool) {
	open := lastUnclosedParen(prefix)
	if open < 0 {
		if call, ok := parenlessCallOnLine(prefix); ok {
			activePos := max(0, len(call.Arguments)-1)
			activeName := ""
			if len(call.Arguments) > 0 {
				activeName = call.Arguments[len(call.Arguments)-1].Name
			}
			return callContext{Target: call.Target, Arguments: call.Arguments, ActiveName: activeName, ActivePos: activePos}, true
		}
		return callContext{}, false
	}
	if isDeclarationCallPrefix(prefix[:open]) {
		return callContext{}, false
	}
	target := callTargetBeforeOpen(prefix[:open])
	if target == "" {
		return callContext{}, false
	}
	args := parseArguments(prefix[open+1:])
	activePos := max(0, len(args)-1)
	activeName := ""
	if len(args) > 0 {
		activeName = args[len(args)-1].Name
	}
	return callContext{Target: target, Arguments: args, ActiveName: activeName, ActivePos: activePos}, true
}

func (c callContext) ActiveParameter(params []Parameter) int {
	if len(params) == 0 {
		return 0
	}
	if c.ActiveName != "" {
		for i, param := range params {
			if strings.EqualFold(param.Name, c.ActiveName) {
				return i
			}
		}
	}
	if c.ActivePos >= len(params) {
		return len(params) - 1
	}
	return c.ActivePos
}

func (a Analyzer) resolveCallSignature(doc Document, target string, pos Position, open []Document) (Signature, bool, error) {
	target = strings.TrimSpace(target)
	if target == "" {
		return Signature{}, false, nil
	}
	if receiver, memberName, ok := splitCallTarget(target); ok {
		receiverType, ok := a.resolveDocumentExpressionTypeAt(doc, receiver, byteOffsetForPosition(doc.Source, pos))
		if ok {
			if strings.EqualFold(receiverType, "Object") && sheetsDefaultExpression(receiver) {
				receiverType = "Excel.Worksheet"
			}
			if member, found := a.DB.ResolveMember(receiverType, memberName); found {
				return a.signatureFromMember(receiverType, member, a.memberKind(receiverType, memberName)), true, nil
			}
		}
		return Signature{}, false, nil
	}
	if member, found := a.DB.ResolveMember("Excel.Application", target); found {
		return a.signatureFromMember("Excel.Application", member, a.memberKind("Excel.Application", target)), true, nil
	}
	if member, found := a.DB.ResolveMember("VBA.Global", target); found {
		return a.signatureFromMember("VBA.Global", member, a.memberKind("VBA.Global", target)), true, nil
	}
	syms, err := a.WorkspaceSymbols(open, target)
	if err != nil {
		return Signature{}, false, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	for _, sym := range syms {
		if !strings.EqualFold(sym.Name, target) || !callableCompletionSymbol(sym) || !a.visibleCompletionSymbol(doc, currentProcedure, sym) {
			continue
		}
		return signatureFromSymbol(sym), true, nil
	}
	return Signature{}, false, nil
}

func splitCallTarget(target string) (receiver string, member string, ok bool) {
	parts := splitMemberExpression(target)
	if len(parts) < 2 {
		return "", "", false
	}
	member = strings.TrimSpace(parts[len(parts)-1])
	receiver = strings.TrimSpace(strings.TrimSuffix(target, "."+member))
	return receiver, member, receiver != "" && member != ""
}

func (a Analyzer) signatureFromMember(receiverType string, member vbadb.MemberInfo, kind string) Signature {
	if dbType, ok := a.DB.ResolveType(receiverType); ok {
		receiverType = dbType.Name
	}
	params := parametersFromDB(member.Parameters)
	label := memberSignature(receiverType, member, kind)
	return Signature{Label: label, Parameters: params, Documentation: member.Summary}
}

func parametersFromDB(params []vbadb.ParamInfo) []Parameter {
	if len(params) == 0 {
		return nil
	}
	out := make([]Parameter, 0, len(params))
	for _, param := range params {
		out = append(out, Parameter{Name: param.Name, Type: firstNonEmpty(param.Type, "Variant"), Optional: param.Optional})
	}
	return out
}

func signatureFromSymbol(sym Symbol) Signature {
	params := sym.Parameters
	label := strings.TrimSpace(sym.Detail)
	if label == "" || !strings.Contains(label, "(") {
		label = symbolSignatureLabel(sym)
	}
	return Signature{Label: label, Parameters: params}
}

func (a Analyzer) argumentDiagnostics(doc Document) []Diagnostic {
	var out []Diagnostic
	lines := normalizedLines(doc.Source)
	for lineNo, line := range lines {
		code := stripLineComment(line)
		for _, call := range callsOnLine(code) {
			sig, ok, err := a.resolveCallSignature(doc, call.Target, Position{Line: lineNo, Character: utf16Len(line)}, []Document{doc})
			if err != nil || !ok || len(sig.Parameters) == 0 {
				continue
			}
			out = append(out, diagnosticsForCallArguments(lineNo, call, sig)...)
		}
	}
	return out
}

type parsedCall struct {
	Target    string
	Arguments []argument
	Line      string
	Start     int
	End       int
}

func callsOnLine(line string) []parsedCall {
	if isDeclarationCallPrefix(line) {
		return nil
	}
	var out []parsedCall
	out = append(out, parenCallsOnLine(line)...)
	if call, ok := parenlessCallOnLine(line); ok {
		out = append(out, call)
	}
	return out
}

func isDeclarationCallPrefix(text string) bool {
	trimmed := strings.TrimSpace(text)
	lower := strings.ToLower(trimmed)
	for _, prefix := range []string{"sub ", "public sub ", "private sub ", "friend sub ", "function ", "public function ", "private function ", "friend function ", "property get ", "property let ", "property set ", "public property ", "private property ", "friend property ", "declare "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func parenCallsOnLine(line string) []parsedCall {
	var out []parsedCall
	inString := false
	for i := 0; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if inString {
				continue
			}
			target := callTargetBeforeOpen(line[:i])
			if target == "" {
				continue
			}
			close := matchingParen(line, i)
			if close < 0 {
				continue
			}
			out = append(out, parsedCall{
				Target:    target,
				Arguments: parseArguments(line[i+1 : close]),
				Line:      line,
				Start:     max(0, i-len(target)),
				End:       close + 1,
			})
			i = close
		}
	}
	return out
}

func parenlessCallOnLine(line string) (parsedCall, bool) {
	trimmed := strings.TrimSpace(line)
	if trimmed == "" {
		return parsedCall{}, false
	}
	lower := strings.ToLower(trimmed)
	for _, prefix := range []string{"if ", "elseif ", "do ", "loop ", "for ", "dim ", "set ", "let ", "with ", "select ", "case ", "option ", "end ", "debug.print"} {
		if strings.HasPrefix(lower, prefix) {
			return parsedCall{}, false
		}
	}
	re := regexp.MustCompile(`(?i)(?:^|[:\s])([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)+|[A-Za-z_][A-Za-z0-9_]*)(\s+)(.*)$`)
	m := re.FindStringSubmatchIndex(line)
	if len(m) == 0 {
		return parsedCall{}, false
	}
	target := strings.TrimSpace(line[m[2]:m[3]])
	argsText := strings.TrimSpace(line[m[6]:m[7]])
	if target == "" || strings.Contains(argsText, "=") && !strings.Contains(argsText, ":=") {
		return parsedCall{}, false
	}
	return parsedCall{Target: target, Arguments: parseArguments(argsText), Line: line, Start: m[2], End: m[7]}, true
}

func matchingParen(line string, open int) int {
	inString := false
	depth := 0
	for i := open; i < len(line); i++ {
		switch line[i] {
		case '"':
			if inString && i+1 < len(line) && line[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString {
				depth--
				if depth == 0 {
					return i
				}
			}
		}
	}
	return -1
}

func diagnosticsForCallArguments(lineNo int, call parsedCall, sig Signature) []Diagnostic {
	var out []Diagnostic
	minArgs, maxArgs := signatureArity(sig.Parameters)
	got := len(call.Arguments)
	if got < minArgs {
		out = append(out, callDiagnostic(lineNo, call, fmt.Sprintf("Argument count mismatch: %s expects at least %d argument(s), got %d.", sigLabelName(sig.Label), minArgs, got)))
	}
	if maxArgs >= 0 && got > maxArgs {
		out = append(out, callDiagnostic(lineNo, call, fmt.Sprintf("Argument count mismatch: %s expects at most %d argument(s), got %d.", sigLabelName(sig.Label), maxArgs, got)))
	}
	for _, arg := range call.Arguments {
		if arg.Name == "" {
			continue
		}
		if !signatureHasParameter(sig.Parameters, arg.Name) {
			out = append(out, callDiagnostic(lineNo, call, fmt.Sprintf("Unknown named argument: %s.", arg.Name)))
		}
	}
	return out
}

func signatureArity(params []Parameter) (min int, max int) {
	max = len(params)
	for _, param := range params {
		if !param.Optional {
			min++
		}
	}
	return min, max
}

func signatureHasParameter(params []Parameter, name string) bool {
	for _, param := range params {
		if strings.EqualFold(param.Name, name) {
			return true
		}
	}
	return false
}

func callDiagnostic(lineNo int, call parsedCall, msg string) Diagnostic {
	return Diagnostic{
		Code:     "VB030",
		Severity: "warning",
		Source:   "xlflow",
		Message:  msg,
		Range: Range{
			Start: Position{Line: lineNo, Character: utf16Len(call.Line[:max(0, min(call.Start, len(call.Line)))])},
			End:   Position{Line: lineNo, Character: utf16Len(call.Line[:max(0, min(call.End, len(call.Line)))])},
		},
	}
}

func sigLabelName(label string) string {
	name := strings.TrimSpace(label)
	if idx := strings.LastIndex(name, "."); idx >= 0 {
		name = name[idx+1:]
	}
	if idx := strings.Index(name, "("); idx >= 0 {
		name = name[:idx]
	}
	return strings.TrimSpace(name)
}

func symbolSignatureLabel(sym Symbol) string {
	var b strings.Builder
	b.WriteString(sym.Name)
	b.WriteString("(")
	for i, param := range sym.Parameters {
		if i > 0 {
			b.WriteString(", ")
		}
		writeParameterLabel(&b, param)
	}
	b.WriteString(")")
	if sym.ReturnType != "" {
		b.WriteString(" As ")
		b.WriteString(sym.ReturnType)
	}
	return b.String()
}

func writeParameterLabel(b *strings.Builder, param Parameter) {
	if param.Optional {
		b.WriteString("Optional ")
	}
	b.WriteString(param.Name)
	if param.Type != "" {
		b.WriteString(" As ")
		b.WriteString(param.Type)
	}
}

func lastUnclosedParen(prefix string) int {
	inString := false
	var stack []int
	for i := 0; i < len(prefix); i++ {
		switch prefix[i] {
		case '"':
			if inString && i+1 < len(prefix) && prefix[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				stack = append(stack, i)
			}
		case ')':
			if !inString && len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
	if len(stack) == 0 {
		return -1
	}
	return stack[len(stack)-1]
}

func callTargetBeforeOpen(prefix string) string {
	target := expressionBefore(prefix)
	target = strings.TrimSpace(target)
	if strings.HasPrefix(strings.ToLower(target), "call ") {
		target = strings.TrimSpace(target[len("call "):])
	}
	fields := strings.Fields(target)
	if len(fields) > 1 {
		target = fields[len(fields)-1]
	}
	return strings.TrimSpace(target)
}

func parseArguments(text string) []argument {
	if strings.TrimSpace(text) == "" {
		return nil
	}
	parts := splitTopLevel(text, ',')
	args := make([]argument, 0, len(parts))
	for _, part := range parts {
		argText := strings.TrimSpace(part)
		arg := argument{Text: argText}
		if name, value, ok := strings.Cut(argText, ":="); ok && isIdentifier(strings.TrimSpace(name)) {
			arg.Name = strings.TrimSpace(name)
			arg.Text = strings.TrimSpace(value)
		}
		args = append(args, arg)
	}
	return args
}

func splitTopLevel(text string, sep byte) []string {
	inString := false
	depth := 0
	start := 0
	var out []string
	for i := 0; i < len(text); i++ {
		switch text[i] {
		case '"':
			if inString && i+1 < len(text) && text[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
		case '(':
			if !inString {
				depth++
			}
		case ')':
			if !inString && depth > 0 {
				depth--
			}
		default:
			if text[i] == sep && !inString && depth == 0 {
				out = append(out, text[start:i])
				start = i + 1
			}
		}
	}
	out = append(out, text[start:])
	return out
}

func isIdentifier(value string) bool {
	if value == "" {
		return false
	}
	for i, r := range value {
		if i == 0 {
			if !isIdentStartRune(r) {
				return false
			}
			continue
		}
		if !isIdentRune(r) {
			return false
		}
	}
	return true
}

func (a Analyzer) typeCompletions(prefix string, replaceRange Range, doc Document, open []Document) ([]Completion, error) {
	var out []Completion
	for _, typ := range a.typeCompletionNames() {
		if !completionMatches(typ.Label, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        typ.Label,
			Kind:         "type",
			Detail:       typ.Detail,
			ReplaceRange: &replace,
		})
	}
	syms, err := a.WorkspaceSymbols(open, prefix)
	if err != nil {
		return nil, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, replaceRange.End)
	for _, sym := range syms {
		if !a.visibleCompletionSymbol(doc, currentProcedure, sym) || !typeSymbolCompletion(sym) || !completionMatches(sym.Name, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        sym.Name,
			Kind:         completionKindForSymbol(sym.Kind),
			Detail:       sym.Detail,
			ReplaceRange: &replace,
		})
	}
	return uniqueCompletions(out), nil
}

func (a Analyzer) progIDCompletions(prefix string, replaceRange Range) []Completion {
	var out []Completion
	for _, progID := range a.DB.ProgIDsList() {
		if !completionMatches(progID, prefix) {
			continue
		}
		replace := replaceRange
		detail := "ProgID"
		if typ, ok := a.DB.ResolveProgID(progID); ok {
			detail = typ.Name
		}
		out = append(out, Completion{
			Label:        progID,
			Kind:         "type",
			Detail:       detail,
			InsertText:   progID,
			ReplaceRange: &replace,
		})
	}
	return uniqueCompletions(out)
}

func (a Analyzer) callCompletions(prefix string, replaceRange Range, doc Document, pos Position, open []Document) ([]Completion, error) {
	syms, err := a.WorkspaceSymbols(open, prefix)
	if err != nil {
		return nil, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	var out []Completion
	for _, sym := range syms {
		if !callableCompletionSymbol(sym) || !a.visibleCompletionSymbol(doc, currentProcedure, sym) || !completionMatches(sym.Name, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        sym.Name,
			Kind:         completionKindForSymbol(sym.Kind),
			Detail:       sym.Detail,
			ReplaceRange: &replace,
		})
	}
	return uniqueCompletions(out), nil
}

func (a Analyzer) setRHSCompletions(prefix string, replaceRange Range, doc Document, pos Position, open []Document) ([]Completion, error) {
	var out []Completion
	replace := replaceRange
	if completionMatches("CreateObject", prefix) {
		out = append(out, Completion{
			Label:        "CreateObject",
			Kind:         "function",
			Detail:       "Create object from ProgID",
			InsertText:   "CreateObject(\"${1:Scripting.Dictionary}\")",
			Snippet:      true,
			ReplaceRange: &replace,
		})
	}
	for _, global := range a.DB.GlobalsList() {
		if !completionMatches(global.Name, prefix) || !objectLikeType(global.ReturnType) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        global.Name,
			Kind:         "variable",
			Detail:       global.ReturnType,
			ReplaceRange: &replace,
		})
	}
	syms, err := a.WorkspaceSymbols(open, prefix)
	if err != nil {
		return nil, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	for _, sym := range syms {
		if !a.visibleCompletionSymbol(doc, currentProcedure, sym) || !completionMatches(sym.Name, prefix) {
			continue
		}
		if !setRHSCompletionSymbol(sym) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        sym.Name,
			Kind:         completionKindForSymbol(sym.Kind),
			Detail:       sym.Detail,
			ReplaceRange: &replace,
		})
	}
	return uniqueCompletions(out), nil
}

func (a Analyzer) valueRHSCompletions(prefix string, replaceRange Range, doc Document, pos Position, open []Document) ([]Completion, error) {
	var out []Completion
	for _, constant := range a.DB.ConstantsList() {
		if !completionMatches(constant.Name, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:         constant.Name,
			Kind:          "constant",
			Detail:        firstNonEmpty(constant.EnumGroup, constant.Type, constant.Kind),
			Documentation: constant.Summary,
			ReplaceRange:  &replace,
		})
	}
	for _, global := range a.DB.GlobalsList() {
		if !completionMatches(global.Name, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        global.Name,
			Kind:         "variable",
			Detail:       global.ReturnType,
			ReplaceRange: &replace,
		})
	}
	for _, control := range a.formControls(doc) {
		if !completionMatches(control.Name, prefix) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        control.Name,
			Kind:         "variable",
			Detail:       control.Type,
			ReplaceRange: &replace,
		})
	}
	out = append(out, localValueCompletionsFromSource(doc, pos, prefix, replaceRange)...)
	syms, err := a.WorkspaceSymbols(open, prefix)
	if err != nil {
		return nil, err
	}
	currentProcedure := currentProcedureNameForDocument(doc, pos)
	for _, sym := range syms {
		if !a.visibleCompletionSymbol(doc, currentProcedure, sym) || !completionMatches(sym.Name, prefix) {
			continue
		}
		if !valueRHSCompletionSymbol(sym) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:        sym.Name,
			Kind:         completionKindForSymbol(sym.Kind),
			Detail:       sym.Detail,
			ReplaceRange: &replace,
		})
	}
	return uniqueCompletions(out), nil
}

func localValueCompletionsFromSource(doc Document, pos Position, prefix string, replaceRange Range) []Completion {
	lines := normalizedLines(doc.Source)
	if pos.Line >= len(lines) {
		return nil
	}
	startLine := 0
	if scope, ok := currentProcedureRangeForDocument(doc, pos); ok {
		startLine = scope.Start.Line
	}
	var out []Completion
	dimRe := regexp.MustCompile(`(?i)^\s*Dim\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?`)
	constRe := regexp.MustCompile(`(?i)^\s*(?:Private\s+|Public\s+)?Const\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:As\s+([A-Za-z_][A-Za-z0-9_.]*))?`)
	for lineNo := startLine; lineNo <= pos.Line && lineNo < len(lines); lineNo++ {
		text := stripLineComment(lines[lineNo])
		if m := dimRe.FindStringSubmatch(text); len(m) > 0 {
			if completionMatches(m[1], prefix) {
				replace := replaceRange
				out = append(out, Completion{
					Label:        m[1],
					Kind:         "variable",
					Detail:       firstNonEmpty(m[2], "Variant"),
					ReplaceRange: &replace,
				})
			}
		}
		if m := constRe.FindStringSubmatch(text); len(m) > 0 {
			if completionMatches(m[1], prefix) {
				replace := replaceRange
				out = append(out, Completion{
					Label:        m[1],
					Kind:         "constant",
					Detail:       firstNonEmpty(m[2], "Variant"),
					ReplaceRange: &replace,
				})
			}
		}
	}
	return out
}

type typeCompletionName struct {
	Label  string
	Detail string
}

func (a Analyzer) typeCompletionNames() []typeCompletionName {
	seen := map[string]bool{}
	out := make([]typeCompletionName, 0)
	add := func(label, detail string) {
		label = strings.TrimSpace(label)
		if label == "" {
			return
		}
		key := strings.ToLower(label)
		if seen[key] {
			return
		}
		seen[key] = true
		out = append(out, typeCompletionName{Label: label, Detail: detail})
	}
	for _, name := range builtinVBATypeNames {
		add(name, "VBA built-in type")
	}
	for _, name := range a.DB.TypeNames() {
		add(name, "type")
	}
	for _, canonical := range a.DB.Aliases {
		typ, ok := a.DB.ResolveType(canonical)
		if !ok {
			continue
		}
		for _, label := range append([]string{shortTypeName(typ.Name)}, typ.Aliases...) {
			detail := typ.Name
			if strings.EqualFold(label, typ.Name) {
				detail = "type"
			}
			add(label, detail)
		}
	}
	sort.SliceStable(out, func(i, j int) bool {
		return strings.ToLower(out[i].Label) < strings.ToLower(out[j].Label)
	})
	return out
}

var builtinVBATypeNames = []string{
	"Boolean",
	"Byte",
	"Currency",
	"Date",
	"Decimal",
	"Double",
	"Integer",
	"Long",
	"LongLong",
	"LongPtr",
	"Object",
	"Single",
	"String",
	"Variant",
}

func shortTypeName(name string) string {
	if idx := strings.LastIndex(name, "."); idx >= 0 {
		return name[idx+1:]
	}
	return name
}

func typeSymbolCompletion(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "type", "enum":
		return true
	case "module":
		return strings.EqualFold(sym.ModuleKind, "class")
	default:
		return false
	}
}

func (a Analyzer) visibleCompletionSymbol(doc Document, currentProcedure string, sym Symbol) bool {
	if strings.EqualFold(sym.Kind, "module") {
		return true
	}
	if sym.Parent != "" {
		return currentProcedure != "" && strings.EqualFold(sym.Parent, currentProcedure) && a.sameDocumentSymbol(doc, sym)
	}
	if a.sameDocumentSymbol(doc, sym) {
		return true
	}
	return !strings.EqualFold(sym.Visibility, "Private")
}

func currentProcedureNameForDocument(doc Document, pos Position) string {
	name, _ := currentProcedureForDocument(doc, pos)
	return name
}

func currentProcedureRangeForDocument(doc Document, pos Position) (Range, bool) {
	_, scope := currentProcedureForDocument(doc, pos)
	if scope == nil {
		return Range{}, false
	}
	return *scope, true
}

func currentProcedureForDocument(doc Document, pos Position) (string, *Range) {
	lines := normalizedLines(doc.Source)
	depth := 0
	current := ""
	var scope *Range
	for lineNo, line := range lines {
		if lineNo > pos.Line && scope == nil {
			break
		}
		text := strings.TrimSpace(line[:codeLimit(line)])
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		switch {
		case strings.HasPrefix(lower, "end sub") || strings.HasPrefix(lower, "end function") || strings.HasPrefix(lower, "end property"):
			if depth > 0 {
				depth--
			}
			if depth == 0 {
				if scope != nil {
					scope.End = Position{Line: lineNo, Character: utf16Len(line)}
				}
				if lineNo < pos.Line {
					current = ""
					scope = nil
				}
				if lineNo >= pos.Line {
					return current, scope
				}
			}
		case procedureStartLine(lower):
			if lineNo > pos.Line {
				return current, scope
			}
			depth++
			if depth == 1 {
				current = procedureNameFromLine(text)
				scope = &Range{Start: Position{Line: lineNo, Character: 0}, End: Position{Line: len(lines), Character: 0}}
			}
		}
	}
	return current, scope
}

func procedureNameFromLine(text string) string {
	fields := strings.Fields(text)
	for len(fields) > 0 {
		switch strings.ToLower(fields[0]) {
		case "public", "private", "friend", "static":
			fields = fields[1:]
		default:
			goto done
		}
	}
done:
	if len(fields) == 0 {
		return ""
	}
	switch strings.ToLower(fields[0]) {
	case "sub", "function":
		if len(fields) > 1 {
			return trimProcedureName(fields[1])
		}
	case "property":
		if len(fields) > 2 {
			return trimProcedureName(fields[2])
		}
	}
	return ""
}

func trimProcedureName(name string) string {
	if idx := strings.IndexAny(name, "("); idx >= 0 {
		name = name[:idx]
	}
	return strings.TrimSpace(name)
}

func (a Analyzer) sameDocumentSymbol(doc Document, sym Symbol) bool {
	if doc.URI != "" && sym.File != "" && sym.File == doc.URI {
		return true
	}
	docKeys := a.workspacePathKeys(doc.Path)
	if len(docKeys) == 0 {
		return false
	}
	return hasAnyPathKey(keySet(docKeys), a.workspacePathKeys(sym.File))
}

func keySet(keys []string) map[string]bool {
	out := make(map[string]bool, len(keys))
	for _, key := range keys {
		out[key] = true
	}
	return out
}

func (a Analyzer) inferWordType(doc Document, word string) (string, bool) {
	return a.inferWordTypeAt(doc, word, -1)
}

type inferredType struct {
	Type   string
	Source string
}

func (a Analyzer) inferWordTypeAt(doc Document, word string, offset int) (string, bool) {
	inferred, ok := a.inferWordTypeInfoAt(doc, word, offset)
	return inferred.Type, ok
}

func (a Analyzer) inferWordTypeInfoAt(doc Document, word string, offset int) (inferredType, bool) {
	if strings.EqualFold(word, "Me") && a.isFormDocument(doc) {
		return inferredType{Type: "MSForms.UserForm", Source: "UserForm instance"}, true
	}
	if control, ok := a.resolveFormControl(doc, word); ok {
		return inferredType{Type: control.Type, Source: "UserForm control"}, true
	}
	if typ, ok := a.DB.ResolveGlobal(word); ok {
		return inferredType{Type: typ.Name, Source: "built-in global"}, true
	}
	var declared string
	declRe := regexp.MustCompile(`(?i)\b` + regexp.QuoteMeta(word) + `\b(?:\s*\([^)]*\))?\s+As\s+(?:New\s+)?([A-Za-z_][A-Za-z0-9_.]*)`)
	if typ, ok := bestTypeMatch(doc.Source, declRe, offset, 1); ok {
		declared = typ
		if !isObjectFallbackType(declared) {
			return inferredType{Type: declared, Source: "declaration"}, true
		}
	}
	newRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*New\s+([A-Za-z_][A-Za-z0-9_.]*)`)
	if typ, ok := bestTypeMatch(doc.Source, newRe, offset, 1); ok {
		return inferredType{Type: typ, Source: "inferred from Set New"}, true
	}
	createRe := regexp.MustCompile(`(?i)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*CreateObject\s*\(\s*"([^"]+)"\s*\)`)
	if progID, ok := bestTypeMatch(doc.Source, createRe, offset, 1); ok {
		if typ, ok := a.DB.ResolveProgID(progID); ok {
			return inferredType{Type: typ.Name, Source: "inferred from CreateObject"}, true
		}
	}
	if expr, exprOffset, ok := bestSetAssignmentExpression(doc.Source, word, offset); ok {
		if typ, ok := a.resolveDocumentExpressionTypeAt(doc, expr, exprOffset); ok {
			return inferredType{Type: typ, Source: "inferred from Set assignment"}, true
		}
	}
	if declared != "" {
		return inferredType{Type: declared, Source: "declaration"}, true
	}
	return inferredType{}, false
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

func (a Analyzer) memberHover(doc Document, word string, wordRange Range, offset int) (*Hover, bool) {
	line := lineAt(doc.Source, wordRange.Start.Line)
	if line == "" {
		return nil, false
	}
	startByte := byteIndexForUTF16(line, wordRange.Start.Character)
	if startByte > len(line) {
		startByte = len(line)
	}
	beforeWord := strings.TrimRight(line[:startByte], " \t")
	if !strings.HasSuffix(beforeWord, ".") {
		return nil, false
	}
	beforeDot := strings.TrimSuffix(beforeWord, ".")
	receiverExpr := expressionBefore(beforeDot)
	if withRelativeDotPrefix(beforeDot) {
		receiverExpr = ""
	}
	if receiverExpr == "" {
		receiverType, ok := a.withBlockTypeAt(doc, wordRange.Start, offset)
		if !ok {
			return nil, false
		}
		if typ, ok := a.DB.ResolveType(receiverType); ok {
			receiverType = typ.Name
		}
		member, ok := a.DB.ResolveMember(receiverType, word)
		if !ok {
			return nil, false
		}
		return &Hover{Contents: memberHover(receiverType, member, a.memberKind(receiverType, word)), Range: wordRange}, true
	}
	if strings.EqualFold(receiverExpr, "Me") {
		if control, ok := a.resolveFormControl(doc, word); ok {
			return &Hover{Contents: variableHover(word, control.Type, "UserForm control"), Range: wordRange}, true
		}
	}
	receiverType, ok := a.resolveDocumentExpressionTypeAt(doc, receiverExpr, offset)
	if !ok {
		return nil, false
	}
	if typ, ok := a.DB.ResolveType(receiverType); ok {
		receiverType = typ.Name
	}
	if strings.EqualFold(receiverType, "Object") && sheetsDefaultExpression(receiverExpr) {
		if _, ok := a.DB.ResolveMember("Excel.Worksheet", word); ok {
			receiverType = "Excel.Worksheet"
		}
	}
	member, ok := a.DB.ResolveMember(receiverType, word)
	if !ok {
		return nil, false
	}
	return &Hover{Contents: memberHover(receiverType, member, a.memberKind(receiverType, word)), Range: wordRange}, true
}

func (a Analyzer) memberKind(receiverType, memberName string) string {
	typ, ok := a.DB.ResolveType(receiverType)
	if !ok {
		return ""
	}
	for _, member := range typ.Properties {
		if strings.EqualFold(member.Name, memberName) {
			return "property"
		}
	}
	for _, member := range typ.Methods {
		if strings.EqualFold(member.Name, memberName) {
			return "method"
		}
	}
	for _, member := range typ.Events {
		if strings.EqualFold(member.Name, memberName) {
			return "event"
		}
	}
	if typ.DefaultMember != "" && strings.EqualFold(typ.DefaultMember, memberName) {
		return "default_member"
	}
	return ""
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
		if strings.EqualFold(current, "Object") && sheetsDefaultExpression(parts[0]) {
			current = "Excel.Worksheet"
		}
	}
	pendingSheetsDefault := strings.EqualFold(current, "Object") && sheetsDefaultExpression(parts[0])
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
		if pendingSheetsDefault {
			if _, ok := a.DB.ResolveMember("Excel.Worksheet", member); ok {
				current = "Excel.Worksheet"
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
			if strings.EqualFold(current, "Object") && strings.EqualFold(member, "Sheets") {
				current = "Excel.Worksheet"
			}
		}
		pendingSheetsDefault = called && strings.EqualFold(current, "Object") && strings.EqualFold(member, "Sheets")
	}
	return current, true
}

func sheetsDefaultExpression(expr string) bool {
	parts := splitMemberExpression(expr)
	if len(parts) == 0 {
		return false
	}
	last := strings.TrimSpace(parts[len(parts)-1])
	if !strings.Contains(last, "(") {
		return false
	}
	if idx := strings.Index(last, "("); idx >= 0 {
		last = strings.TrimSpace(last[:idx])
	}
	return strings.EqualFold(last, "Sheets")
}

func (a Analyzer) withBlockTypeAt(doc Document, pos Position, offset int) (string, bool) {
	lines := normalizedLines(doc.Source)
	if pos.Line <= 0 || pos.Line > len(lines) {
		return "", false
	}
	var stack []string
	for lineNo := 0; lineNo < pos.Line; lineNo++ {
		trimmed := strings.TrimSpace(stripLineComment(lines[lineNo]))
		if trimmed == "" {
			continue
		}
		if regexp.MustCompile(`(?i)^End\s+With\b`).MatchString(trimmed) {
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
			continue
		}
		m := regexp.MustCompile(`(?i)^With\s+(.+)$`).FindStringSubmatch(trimmed)
		if len(m) == 0 {
			continue
		}
		if typ, ok := a.resolveWithExpressionTypeAt(doc, strings.TrimSpace(m[1]), stack, offset); ok {
			stack = append(stack, typ)
		} else {
			stack = append(stack, "")
		}
	}
	for i := len(stack) - 1; i >= 0; i-- {
		if stack[i] != "" {
			return stack[i], true
		}
	}
	return "", false
}

func (a Analyzer) resolveWithExpressionTypeAt(doc Document, expr string, stack []string, offset int) (string, bool) {
	expr = strings.TrimSpace(expr)
	if strings.HasPrefix(expr, ".") {
		if len(stack) == 0 || stack[len(stack)-1] == "" {
			return "", false
		}
		return a.resolveMemberChainFromType(stack[len(stack)-1], expr)
	}
	return a.resolveDocumentExpressionTypeAt(doc, expr, offset)
}

func (a Analyzer) resolveMemberChainFromType(baseType, expr string) (string, bool) {
	current := baseType
	for _, raw := range splitMemberExpression(strings.TrimPrefix(strings.TrimSpace(expr), ".")) {
		member := strings.TrimSpace(raw)
		called := strings.Contains(member, "(")
		if idx := strings.Index(member, "("); idx >= 0 {
			member = strings.TrimSpace(member[:idx])
		}
		if member == "" {
			continue
		}
		info, ok := a.DB.ResolveMember(current, member)
		if !ok || info.ReturnType == "" {
			return "", false
		}
		current = info.ReturnType
		if called {
			if typ, ok := a.collectionDefaultType(current); ok {
				current = typ
			}
		}
	}
	return current, true
}

func stripLineComment(line string) string {
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
				return line[:i]
			}
		}
	}
	return line
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
	return bestType, bestType != ""
}

func bestSetAssignmentExpression(source, word string, offset int) (string, int, bool) {
	re := regexp.MustCompile(`(?im)\bSet\s+` + regexp.QuoteMeta(word) + `\s*=\s*([^\r\n:]+)`)
	matches := re.FindAllStringSubmatchIndex(source, -1)
	bestStart := -1
	bestExpr := ""
	bestExprOffset := -1
	for _, match := range matches {
		if len(match) < 4 || match[2] < 0 || match[3] < 0 {
			continue
		}
		start := match[0]
		if offset >= 0 && start > offset {
			continue
		}
		expr := strings.TrimSpace(stripLineComment(source[match[2]:match[3]]))
		if expr == "" {
			continue
		}
		if bestStart < 0 || start > bestStart {
			bestStart = start
			bestExpr = expr
			bestExprOffset = match[2]
		}
	}
	return bestExpr, bestExprOffset, bestExpr != ""
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
			Detail:        a.signatureFromMember(receiverType, member, kind).Label,
			Documentation: member.Summary,
		})
	}
	return uniqueCompletions(out)
}

func (a Analyzer) moduleMemberCompletions(open []Document, moduleName, prefix string) ([]Completion, error) {
	moduleName = strings.TrimSpace(moduleName)
	if moduleName == "" {
		return nil, nil
	}
	syms, err := a.WorkspaceSymbols(open, "")
	if err != nil {
		return nil, err
	}
	var out []Completion
	for _, sym := range syms {
		if !strings.EqualFold(sym.Module, moduleName) || !moduleMemberCompletionSymbol(sym) {
			continue
		}
		if !completionMatches(sym.Name, prefix) {
			continue
		}
		out = append(out, Completion{
			Label:  sym.Name,
			Kind:   completionKindForSymbol(sym.Kind),
			Detail: sym.Detail,
		})
	}
	return uniqueCompletions(out), nil
}

func moduleMemberCompletionSymbol(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
	default:
		return false
	}
	return !strings.EqualFold(sym.Visibility, "Private")
}

func callableCompletionSymbol(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "sub", "function", "declare_sub", "declare_function":
		return true
	default:
		return false
	}
}

func setRHSCompletionSymbol(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "class":
		return true
	case "module":
		return strings.EqualFold(sym.ModuleKind, "class") || strings.EqualFold(sym.ModuleKind, "form")
	case "function", "property", "property_get":
		return objectLikeType(returnTypeFromDetail(sym.Detail))
	default:
		return false
	}
}

func valueRHSCompletionSymbol(sym Symbol) bool {
	switch strings.ToLower(sym.Kind) {
	case "const", "module_variable", "local_variable", "field", "parameter", "function", "property", "property_get":
		return true
	default:
		return false
	}
}

func returnTypeFromDetail(detail string) string {
	re := regexp.MustCompile(`(?i)\bAs\s+([A-Za-z_][A-Za-z0-9_.]*)\s*$`)
	if m := re.FindStringSubmatch(strings.TrimSpace(detail)); len(m) > 1 {
		return m[1]
	}
	return ""
}

func objectLikeType(name string) bool {
	name = strings.TrimSpace(name)
	if name == "" {
		return false
	}
	if isObjectFallbackType(name) {
		return true
	}
	if strings.Contains(name, ".") {
		return true
	}
	switch strings.ToLower(name) {
	case "boolean", "byte", "currency", "date", "decimal", "double", "integer", "long", "longlong", "longptr", "single", "string", "variant", "any":
		return false
	default:
		return true
	}
}

func (a Analyzer) syntaxCompletions(doc Document, pos Position, prefix string) []Completion {
	start, typed, ok := statementPrefix(prefix)
	if !ok {
		return nil
	}
	replaceRange := Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}
	if isModuleLevelPosition(doc.Source, pos) {
		return completionsFromSpecs(moduleDeclarationCompletions, typed, replaceRange)
	}
	return completionsFromSpecs(procedureStatementCompletions, typed, replaceRange)
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

type syntaxCompletionSpec struct {
	label         string
	detail        string
	insertText    string
	documentation string
}

var moduleDeclarationCompletions = []syntaxCompletionSpec{
	{
		label:         "Option Explicit",
		detail:        "Require explicit variable declarations",
		insertText:    "Option Explicit",
		documentation: "Adds the module-level `Option Explicit` declaration.",
	},
	{
		label:         "Option Base 1",
		detail:        "Set default array lower bound",
		insertText:    "Option Base 1",
		documentation: "Adds the module-level `Option Base 1` declaration.",
	},
	{
		label:         "Option Private Module",
		detail:        "Hide module members from external projects",
		insertText:    "Option Private Module",
		documentation: "Adds the module-level `Option Private Module` declaration.",
	},
	{
		label:         "Public",
		detail:        "Public declaration modifier",
		insertText:    "Public ",
		documentation: "Starts a public module-level declaration.",
	},
	{
		label:         "Private",
		detail:        "Private declaration modifier",
		insertText:    "Private ",
		documentation: "Starts a private module-level declaration.",
	},
	{
		label:         "Friend",
		detail:        "Friend declaration modifier",
		insertText:    "Friend ",
		documentation: "Starts a friend declaration in a class module.",
	},
	{
		label:         "Sub",
		detail:        "Procedure declaration",
		insertText:    "Sub ${1:Name}()\n    $0\nEnd Sub",
		documentation: "Creates a module-level Sub procedure.",
	},
	{
		label:         "Public Sub",
		detail:        "Public procedure declaration",
		insertText:    "Public Sub ${1:Name}()\n    $0\nEnd Sub",
		documentation: "Creates a public Sub procedure.",
	},
	{
		label:         "Private Sub",
		detail:        "Private procedure declaration",
		insertText:    "Private Sub ${1:Name}()\n    $0\nEnd Sub",
		documentation: "Creates a private Sub procedure.",
	},
	{
		label:         "Function",
		detail:        "Function declaration",
		insertText:    "Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function",
		documentation: "Creates a module-level Function procedure.",
	},
	{
		label:         "Public Function",
		detail:        "Public function declaration",
		insertText:    "Public Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function",
		documentation: "Creates a public Function procedure.",
	},
	{
		label:         "Private Function",
		detail:        "Private function declaration",
		insertText:    "Private Function ${1:Name}() As ${2:Variant}\n    $0\nEnd Function",
		documentation: "Creates a private Function procedure.",
	},
	{
		label:         "Property Get",
		detail:        "Property getter declaration",
		insertText:    "Property Get ${1:Name}() As ${2:Variant}\n    $0\nEnd Property",
		documentation: "Creates a Property Get procedure.",
	},
	{
		label:         "Property Let",
		detail:        "Property setter declaration",
		insertText:    "Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})\n    $0\nEnd Property",
		documentation: "Creates a Property Let procedure.",
	},
	{
		label:         "Property Set",
		detail:        "Object property setter declaration",
		insertText:    "Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})\n    $0\nEnd Property",
		documentation: "Creates a Property Set procedure.",
	},
	{
		label:         "Public Property Get",
		detail:        "Public property getter declaration",
		insertText:    "Public Property Get ${1:Name}() As ${2:Variant}\n    $0\nEnd Property",
		documentation: "Creates a public Property Get procedure.",
	},
	{
		label:         "Public Property Let",
		detail:        "Public property setter declaration",
		insertText:    "Public Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})\n    $0\nEnd Property",
		documentation: "Creates a public Property Let procedure.",
	},
	{
		label:         "Public Property Set",
		detail:        "Public object property setter declaration",
		insertText:    "Public Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})\n    $0\nEnd Property",
		documentation: "Creates a public Property Set procedure.",
	},
	{
		label:         "Private Property Get",
		detail:        "Private property getter declaration",
		insertText:    "Private Property Get ${1:Name}() As ${2:Variant}\n    $0\nEnd Property",
		documentation: "Creates a private Property Get procedure.",
	},
	{
		label:         "Private Property Let",
		detail:        "Private property setter declaration",
		insertText:    "Private Property Let ${1:Name}(ByVal ${2:value} As ${3:Variant})\n    $0\nEnd Property",
		documentation: "Creates a private Property Let procedure.",
	},
	{
		label:         "Private Property Set",
		detail:        "Private object property setter declaration",
		insertText:    "Private Property Set ${1:Name}(ByVal ${2:value} As ${3:Object})\n    $0\nEnd Property",
		documentation: "Creates a private Property Set procedure.",
	},
	{
		label:         "Dim",
		detail:        "Module variable declaration",
		insertText:    "Dim ${1:name} As ${2:Variant}",
		documentation: "Declares a module-level variable.",
	},
	{
		label:         "Const",
		detail:        "Constant declaration",
		insertText:    "Const ${1:Name} As ${2:Variant} = ${3:value}",
		documentation: "Declares a module-level constant.",
	},
	{
		label:         "Public Const",
		detail:        "Public constant declaration",
		insertText:    "Public Const ${1:Name} As ${2:Variant} = ${3:value}",
		documentation: "Declares a public module-level constant.",
	},
	{
		label:         "Private Const",
		detail:        "Private constant declaration",
		insertText:    "Private Const ${1:Name} As ${2:Variant} = ${3:value}",
		documentation: "Declares a private module-level constant.",
	},
	{
		label:         "Type",
		detail:        "User-defined type declaration",
		insertText:    "Type ${1:Name}\n    ${2:Field} As ${3:Variant}\nEnd Type",
		documentation: "Declares a user-defined type.",
	},
	{
		label:         "Public Type",
		detail:        "Public user-defined type declaration",
		insertText:    "Public Type ${1:Name}\n    ${2:Field} As ${3:Variant}\nEnd Type",
		documentation: "Declares a public user-defined type.",
	},
	{
		label:         "Private Type",
		detail:        "Private user-defined type declaration",
		insertText:    "Private Type ${1:Name}\n    ${2:Field} As ${3:Variant}\nEnd Type",
		documentation: "Declares a private user-defined type.",
	},
	{
		label:         "Enum",
		detail:        "Enum declaration",
		insertText:    "Enum ${1:Name}\n    ${2:Member} = ${3:0}\nEnd Enum",
		documentation: "Declares an enum.",
	},
	{
		label:         "Public Enum",
		detail:        "Public enum declaration",
		insertText:    "Public Enum ${1:Name}\n    ${2:Member} = ${3:0}\nEnd Enum",
		documentation: "Declares a public enum.",
	},
	{
		label:         "Private Enum",
		detail:        "Private enum declaration",
		insertText:    "Private Enum ${1:Name}\n    ${2:Member} = ${3:0}\nEnd Enum",
		documentation: "Declares a private enum.",
	},
	{
		label:         "Declare PtrSafe Function",
		detail:        "External function declaration",
		insertText:    "Public Declare PtrSafe Function ${1:Name} Lib \"${2:library}\" (${3:args}) As ${4:LongPtr}",
		documentation: "Declares an external PtrSafe function.",
	},
	{
		label:         "Declare PtrSafe Sub",
		detail:        "External sub declaration",
		insertText:    "Public Declare PtrSafe Sub ${1:Name} Lib \"${2:library}\" (${3:args})",
		documentation: "Declares an external PtrSafe subroutine.",
	},
}

var procedureStatementCompletions = []syntaxCompletionSpec{
	{
		label:         "Dim",
		detail:        "Local variable declaration",
		insertText:    "Dim ${1:name} As ${2:Variant}",
		documentation: "Declares a procedure-local variable.",
	},
	{
		label:         "Set",
		detail:        "Object assignment",
		insertText:    "Set ${1:target} = ${2:expression}",
		documentation: "Assigns an object reference in a procedure.",
	},
	{
		label:         "If Then",
		detail:        "Conditional block",
		insertText:    "If ${1:condition} Then\n    $0\nEnd If",
		documentation: "Creates an If block.",
	},
	{
		label:         "If Else",
		detail:        "Conditional block with Else",
		insertText:    "If ${1:condition} Then\n    ${2:statements}\nElse\n    $0\nEnd If",
		documentation: "Creates an If block with an Else branch.",
	},
	{
		label:         "For Next",
		detail:        "Counter loop",
		insertText:    "For ${1:i} = ${2:1} To ${3:10}\n    $0\nNext ${1:i}",
		documentation: "Creates a For...Next loop.",
	},
	{
		label:         "For Each",
		detail:        "Collection loop",
		insertText:    "For Each ${1:item} In ${2:collection}\n    $0\nNext ${1:item}",
		documentation: "Creates a For Each loop.",
	},
	{
		label:         "Do While",
		detail:        "Condition loop",
		insertText:    "Do While ${1:condition}\n    $0\nLoop",
		documentation: "Creates a Do While loop.",
	},
	{
		label:         "Do Until",
		detail:        "Condition loop",
		insertText:    "Do Until ${1:condition}\n    $0\nLoop",
		documentation: "Creates a Do Until loop.",
	},
	{
		label:         "Select Case",
		detail:        "Select block",
		insertText:    "Select Case ${1:expression}\nCase ${2:value}\n    $0\nEnd Select",
		documentation: "Creates a Select Case block.",
	},
	{
		label:         "With",
		detail:        "With block",
		insertText:    "With ${1:expression}\n    $0\nEnd With",
		documentation: "Creates a With block.",
	},
}

func completionsFromSpecs(specs []syntaxCompletionSpec, prefix string, replaceRange Range) []Completion {
	var out []Completion
	for _, spec := range specs {
		if !completionMatches(spec.label, prefix) {
			continue
		}
		if strings.EqualFold(strings.TrimSpace(spec.label), strings.TrimSpace(prefix)) {
			continue
		}
		replace := replaceRange
		out = append(out, Completion{
			Label:         spec.label,
			Kind:          "snippet",
			Detail:        spec.detail,
			Documentation: spec.documentation,
			InsertText:    spec.insertText,
			Snippet:       strings.Contains(spec.insertText, "$"),
			ReplaceRange:  &replace,
		})
	}
	return out
}

func memberCompletionContext(prefix string) (memberPrefix string, receiverExpr string, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !strings.HasSuffix(beforeWord, ".") {
		return "", "", false
	}
	receiver := expressionBefore(strings.TrimSuffix(beforeWord, "."))
	receiver = trimWithExpressionPrefix(receiver)
	if receiver == "" {
		return "", "", false
	}
	return wordPrefix, receiver, true
}

func trimWithExpressionPrefix(expr string) string {
	trimmed := strings.TrimSpace(expr)
	if len(trimmed) < len("With ") || !strings.EqualFold(trimmed[:len("With")], "With") || !isSpaceByte(trimmed[len("With")]) {
		return expr
	}
	rest := strings.TrimSpace(trimmed[len("With"):])
	if rest == "" {
		return expr
	}
	return rest
}

func isSpaceByte(ch byte) bool {
	return ch == ' ' || ch == '\t'
}

func (a Analyzer) withBlockMemberCompletionContext(doc Document, pos Position, prefix string) (memberPrefix string, receiverType string, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !strings.HasSuffix(beforeWord, ".") {
		return "", "", false
	}
	beforeDot := strings.TrimSuffix(beforeWord, ".")
	if strings.TrimSpace(beforeDot) != "" && !withRelativeDotPrefix(beforeDot) {
		return "", "", false
	}
	typ, ok := a.withBlockTypeAt(doc, pos, byteOffsetForPosition(doc.Source, pos))
	if !ok {
		return "", "", false
	}
	return wordPrefix, typ, true
}

func withRelativeDotPrefix(beforeDot string) bool {
	fields := strings.Fields(beforeDot)
	return len(fields) > 0 && strings.EqualFold(fields[len(fields)-1], "With")
}

func typeCompletionContext(prefix string, pos Position) (typePrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !endsWithTypeIntro(beforeWord) {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func callCompletionContext(prefix string, pos Position) (callPrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !regexp.MustCompile(`(?i)(^|:)\s*Call\s*$`).MatchString(beforeWord) {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func setRHSCompletionContext(prefix string, pos Position) (rhsPrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !regexp.MustCompile(`(?i)(^|:)\s*Set\s+[A-Za-z_][A-Za-z0-9_]*\s*=\s*$`).MatchString(beforeWord) {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func valueRHSCompletionContext(prefix string, pos Position) (rhsPrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	eq := assignmentOperatorIndex(beforeWord)
	if eq < 0 {
		return "", Range{}, false
	}
	lhs := strings.TrimSpace(beforeWord[:eq])
	if disallowedValueAssignmentLHS(lhs) {
		return "", Range{}, false
	}
	if strings.TrimSpace(beforeWord[eq+1:]) != "" {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func conditionCompletionContext(prefix string, pos Position) (conditionPrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !conditionExpressionPrefix(beforeWord) {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func forEachInCompletionContext(prefix string, pos Position) (eachPrefix string, replaceRange Range, ok bool) {
	wordPrefix := currentIdentifierPrefix(prefix)
	beforeWord := strings.TrimRight(prefix[:len(prefix)-len(wordPrefix)], " \t")
	if !regexp.MustCompile(`(?i)(^|:)\s*For\s+Each\s+[A-Za-z_][A-Za-z0-9_]*\s+In\s*$`).MatchString(beforeWord) {
		return "", Range{}, false
	}
	start := len(prefix) - len(wordPrefix)
	return wordPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:start])},
		End:   pos,
	}, true
}

func createObjectProgIDCompletionContext(prefix string, pos Position) (progIDPrefix string, replaceRange Range, ok bool) {
	quote := strings.LastIndex(prefix, `"`)
	if quote < 0 {
		return "", Range{}, false
	}
	beforeQuote := prefix[:quote]
	if !regexp.MustCompile(`(?i)\bCreateObject\s*\(\s*$`).MatchString(beforeQuote) {
		return "", Range{}, false
	}
	progIDPrefix = prefix[quote+1:]
	if strings.Contains(progIDPrefix, `"`) {
		return "", Range{}, false
	}
	return progIDPrefix, Range{
		Start: Position{Line: pos.Line, Character: utf16Len(prefix[:quote+1])},
		End:   pos,
	}, true
}

func insideOpenString(prefix string) bool {
	inString := false
	for i := 0; i < len(prefix); i++ {
		if prefix[i] != '"' {
			continue
		}
		if inString && i+1 < len(prefix) && prefix[i+1] == '"' {
			i++
			continue
		}
		inString = !inString
	}
	return inString
}

func assignmentOperatorIndex(prefix string) int {
	inString := false
	depth := 0
	idx := -1
	for i := 0; i < len(prefix); i++ {
		ch := prefix[i]
		if ch == '"' {
			if inString && i+1 < len(prefix) && prefix[i+1] == '"' {
				i++
				continue
			}
			inString = !inString
			continue
		}
		if inString {
			continue
		}
		switch ch {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case '=':
			if depth == 0 && isAssignmentEqual(prefix, i) {
				idx = i
			}
		}
	}
	return idx
}

func isAssignmentEqual(prefix string, idx int) bool {
	if idx > 0 {
		switch prefix[idx-1] {
		case '<', '>', '=':
			return false
		}
	}
	if idx+1 < len(prefix) {
		switch prefix[idx+1] {
		case '<', '>', '=':
			return false
		}
	}
	return true
}

func disallowedValueAssignmentLHS(lhs string) bool {
	fields := strings.Fields(strings.ToLower(lhs))
	if len(fields) == 0 {
		return true
	}
	switch fields[0] {
	case "set", "if", "elseif", "while", "until", "for", "case", "select":
		return true
	case "do", "loop":
		return len(fields) > 1 && (fields[1] == "while" || fields[1] == "until")
	default:
		return false
	}
}

func conditionExpressionPrefix(prefix string) bool {
	clean := strings.TrimSpace(stripLineComment(prefix))
	if clean == "" {
		return false
	}
	lower := strings.ToLower(clean)
	if regexp.MustCompile(`\bthen\b`).MatchString(lower) {
		return false
	}
	switch {
	case lower == "if" || strings.HasPrefix(lower, "if "):
		return true
	case lower == "elseif" || strings.HasPrefix(lower, "elseif "):
		return true
	case lower == "while" || strings.HasPrefix(lower, "while "):
		return true
	case lower == "do while" || strings.HasPrefix(lower, "do while "):
		return true
	case lower == "do until" || strings.HasPrefix(lower, "do until "):
		return true
	case lower == "loop while" || strings.HasPrefix(lower, "loop while "):
		return true
	case lower == "loop until" || strings.HasPrefix(lower, "loop until "):
		return true
	default:
		return false
	}
}

func endsWithTypeIntro(prefix string) bool {
	fields := strings.Fields(prefix)
	if len(fields) == 0 {
		return false
	}
	last := strings.ToLower(fields[len(fields)-1])
	if last == "as" || last == "new" {
		return true
	}
	return false
}

func statementPrefix(prefix string) (start int, typed string, ok bool) {
	start = len(prefix) - len(strings.TrimLeft(prefix, " \t"))
	typed = strings.TrimSpace(prefix[start:])
	if typed == "" {
		return start, typed, true
	}
	for _, r := range typed {
		if isIdentRune(r) || r == ' ' || r == '\t' {
			continue
		}
		return 0, "", false
	}
	return start, strings.Join(strings.Fields(typed), " "), true
}

func isModuleLevelPosition(source string, pos Position) bool {
	lines := normalizedLines(source)
	if pos.Line <= 0 {
		return true
	}
	depth := 0
	for lineNo, line := range lines {
		if lineNo >= pos.Line {
			break
		}
		text := strings.TrimSpace(line[:codeLimit(line)])
		if text == "" {
			continue
		}
		lower := strings.ToLower(text)
		switch {
		case strings.HasPrefix(lower, "end sub") || strings.HasPrefix(lower, "end function") || strings.HasPrefix(lower, "end property"):
			if depth > 0 {
				depth--
			}
		case procedureStartLine(lower):
			depth++
		}
	}
	return depth == 0
}

func procedureStartLine(lower string) bool {
	if strings.HasPrefix(lower, "end ") {
		return false
	}
	fields := strings.Fields(lower)
	for len(fields) > 0 {
		switch fields[0] {
		case "public", "private", "friend", "static":
			fields = fields[1:]
		default:
			goto done
		}
	}
done:
	if len(fields) == 0 {
		return false
	}
	switch fields[0] {
	case "sub", "function":
		return true
	case "property":
		return len(fields) > 1 && (fields[1] == "get" || fields[1] == "let" || fields[1] == "set")
	default:
		return false
	}
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

func completionKindForSymbol(kind string) string {
	switch strings.ToLower(kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return "function"
	case "const":
		return "constant"
	case "module_variable", "local_variable", "field", "parameter":
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
		converted := Symbol{
			Name:       sym.Name,
			Kind:       sym.Kind,
			Detail:     firstNonEmpty(sym.Signature, sym.Kind+" "+sym.Name),
			ReturnType: sym.ReturnType,
			Parameters: symbolParameters(sym.Parameters),
			File:       firstNonEmpty(uri, file.Path, sym.File),
			Module:     sym.Module,
			ModuleKind: file.ModuleKind,
			Parent:     sym.Parent,
			Visibility: sym.Visibility,
			Range: Range{
				Start: Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1)},
				End:   Position{Line: sym.EndLine - 1, Character: max(0, sym.EndColumn-1)},
			},
			Selection: Range{
				Start: Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1)},
				End:   Position{Line: sym.StartLine - 1, Character: max(0, sym.StartColumn-1+utf16Len(sym.Name))},
			},
		}
		if !rangeContains(converted.Range, converted.Selection) {
			converted.Selection = Range{Start: converted.Range.Start, End: converted.Range.Start}
		}
		out = append(out, converted)
	}
	return out
}

func symbolParameters(params []symbols.Parameter) []Parameter {
	if len(params) == 0 {
		return nil
	}
	out := make([]Parameter, 0, len(params))
	for _, param := range params {
		out = append(out, Parameter{
			Name:     param.Name,
			Type:     firstNonEmpty(param.Type, "Variant"),
			Optional: param.Optional,
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

func typeHover(typ vbadb.TypeInfo, source string) string {
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
	if source != "" {
		b.WriteString("\n\nSource: ")
		b.WriteString(source)
	}
	if typ.ElementType != "" {
		b.WriteString("\n\nCollection element: `")
		b.WriteString(typ.ElementType)
		b.WriteString("`")
	}
	return b.String()
}

func variableHover(name, typ, source string) string {
	var b strings.Builder
	b.WriteString("```vb\n")
	b.WriteString(name)
	b.WriteString(" As ")
	b.WriteString(typ)
	b.WriteString("\n```")
	if source != "" {
		b.WriteString("\nSource: ")
		b.WriteString(source)
	}
	return b.String()
}

func symbolHover(signature, source string) string {
	var b strings.Builder
	b.WriteString("```vb\n")
	b.WriteString(signature)
	b.WriteString("\n```")
	if source != "" {
		b.WriteString("\nSource: ")
		b.WriteString(source)
	}
	return b.String()
}

func memberHover(receiverType string, member vbadb.MemberInfo, kind string) string {
	var b strings.Builder
	b.WriteString("```vb\n")
	b.WriteString(memberSignature(receiverType, member, kind))
	b.WriteString("\n```")
	if member.Summary != "" {
		b.WriteString("\n")
		b.WriteString(member.Summary)
	}
	b.WriteString("\n\nSource: built-in ")
	if library, _, ok := strings.Cut(receiverType, "."); ok && library != "" {
		b.WriteString(library)
		b.WriteString(" object model DB")
	} else {
		b.WriteString("VBA/COM type DB")
	}
	return b.String()
}

func memberSignature(receiverType string, member vbadb.MemberInfo, kind string) string {
	var b strings.Builder
	b.WriteString(receiverType)
	b.WriteString(".")
	b.WriteString(member.Name)
	if len(member.Parameters) > 0 {
		b.WriteString("(")
		for i, param := range member.Parameters {
			if i > 0 {
				b.WriteString(", ")
			}
			if param.Optional {
				b.WriteString("Optional ")
			}
			b.WriteString(param.Name)
			if param.Type != "" {
				b.WriteString(" As ")
				b.WriteString(param.Type)
			}
		}
		b.WriteString(")")
	}
	if member.ReturnType != "" {
		b.WriteString(" As ")
		b.WriteString(member.ReturnType)
	} else if kind == "method" {
		b.WriteString(" As void")
	}
	return b.String()
}

func symbolSource(sym Symbol) string {
	if isLocalSymbol(sym) {
		return "declaration"
	}
	switch strings.ToLower(sym.Kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set", "module_variable", "const", "field", "parameter":
		return "declaration"
	case "module", "class":
		return "project symbol"
	default:
		return ""
	}
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

func min(a, b int) int {
	if a < b {
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
