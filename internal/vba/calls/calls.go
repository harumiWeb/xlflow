package calls

import (
	"fmt"
	"path/filepath"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Options struct {
	RootDir string
	Config  config.Config
	Path    string
	From    string
	To      string
}

type Result struct {
	Root    string        `json:"root"`
	Calls   []Call        `json:"calls"`
	Summary ResultSummary `json:"summary"`
}

type ResultSummary struct {
	Files        int `json:"files"`
	Calls        int `json:"calls"`
	Matched      int `json:"matched"`
	Unresolved   int `json:"unresolved"`
	Ambiguous    int `json:"ambiguous"`
	External     int `json:"external"`
	BuiltinLike  int `json:"builtinLike"`
	MemberCalls  int `json:"memberCalls"`
	ParseErrors  int `json:"parseErrors"`
	MissingNodes int `json:"missingNodes"`
}

type Call struct {
	File       string               `json:"file"`
	Module     string               `json:"module"`
	Caller     *Caller              `json:"caller,omitempty"`
	Callee     Callee               `json:"callee"`
	Arguments  Arguments            `json:"arguments"`
	Range      vbaast.Range         `json:"range"`
	Parse      symbols.ParseSummary `json:"parse"`
	Resolution Resolution           `json:"resolution"`
}

type Caller struct {
	Name          string `json:"name"`
	Kind          string `json:"kind"`
	QualifiedName string `json:"qualifiedName"`
}

type Callee struct {
	Text     string  `json:"text"`
	BaseName string  `json:"baseName"`
	Receiver *string `json:"receiver"`
	Member   string  `json:"member"`
}

type Arguments struct {
	Count int             `json:"count"`
	Named []NamedArgument `json:"named"`
}

type NamedArgument struct {
	Name      string `json:"name"`
	ValueText string `json:"valueText"`
}

type Resolution struct {
	Status     string      `json:"status"`
	Candidates []Candidate `json:"candidates,omitempty"`
}

type Candidate struct {
	QualifiedName string `json:"qualifiedName"`
	Kind          string `json:"kind"`
	File          string `json:"file"`
	Line          int    `json:"line"`
}

type extractor struct {
	source     []byte
	file       string
	moduleName string
	parse      symbols.ParseSummary
	current    *Caller
	resolver   resolver
	calls      []Call
}

type resolver struct {
	byName map[string][]Candidate
}

var procedureKinds = map[string]bool{
	"sub":              true,
	"function":         true,
	"property":         true,
	"property_get":     true,
	"property_let":     true,
	"property_set":     true,
	"declare":          true,
	"declare_sub":      true,
	"declare_function": true,
	"event":            true,
}

var builtinLikeNames = map[string]bool{
	"array": true, "asc": true, "cbool": true, "cbyte": true, "ccur": true,
	"cdate": true, "cdbl": true, "cdec": true, "choose": true, "chr": true,
	"cint": true, "clng": true, "clnglng": true, "clngptr": true, "cos": true,
	"createobject": true, "cstr": true, "date": true, "dateadd": true,
	"debug.print": true, "dir": true, "doevents": true, "environ": true,
	"format": true, "getobject": true, "inputbox": true, "instr": true,
	"isarray": true, "isdate": true, "isempty": true, "iserror": true,
	"isnull": true, "isnumeric": true, "join": true, "lbound": true,
	"lcase": true, "left": true, "len": true, "mid": true, "msgbox": true,
	"replace": true, "right": true, "rnd": true, "split": true, "str": true,
	"trim": true, "typename": true, "ubound": true, "ucase": true, "val": true,
}

var externalLikeReceivers = map[string]bool{
	"application": true, "debug": true, "excel": true, "worksheetfunction": true,
}

func Inspect(opts Options) (*Result, error) {
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	absRoot, err := filepath.Abs(rootDir)
	if err != nil {
		return nil, err
	}
	symbolResult, err := symbols.Inspect(symbols.Options{
		RootDir:        absRoot,
		Config:         opts.Config,
		Path:           opts.Path,
		IncludePrivate: true,
		IncludeLabels:  false,
	})
	if err != nil {
		return nil, err
	}
	res := buildResolver(symbolResult)
	parser, err := vbaast.NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()

	result := &Result{Root: symbolResult.Root, Calls: []Call{}}
	for _, file := range symbolResult.Files {
		path := resolveDisplayPath(absRoot, file.Path)
		parsed, err := parser.ParseFile(path)
		if err != nil {
			return nil, err
		}
		ext := extractor{
			source:     parsed.Source,
			file:       file.Path,
			moduleName: file.ModuleName,
			parse: symbols.ParseSummary{
				HasError:   parsed.HasError,
				HasMissing: parsed.HasMissing,
			},
			resolver: res,
		}
		ext.visit(parsed.Root)
		parsed.Close()
		for _, call := range ext.calls {
			if !matchesFrom(call, opts.From) || !matchesTo(call, opts.To) {
				continue
			}
			result.Calls = append(result.Calls, call)
			addResolutionSummary(&result.Summary, call.Resolution.Status)
		}
		if ext.parse.HasError {
			result.Summary.ParseErrors++
		}
		if ext.parse.HasMissing {
			result.Summary.MissingNodes++
		}
	}
	result.Summary.Files = len(symbolResult.Files)
	result.Summary.Calls = len(result.Calls)
	sort.SliceStable(result.Calls, func(i, j int) bool {
		a, b := result.Calls[i], result.Calls[j]
		if a.File != b.File {
			return a.File < b.File
		}
		if a.Range.StartLine != b.Range.StartLine {
			return a.Range.StartLine < b.Range.StartLine
		}
		return a.Range.StartColumn < b.Range.StartColumn
	})
	return result, nil
}

func buildResolver(result *symbols.Result) resolver {
	res := resolver{byName: map[string][]Candidate{}}
	if result == nil {
		return res
	}
	for _, file := range result.Files {
		for _, sym := range file.Symbols {
			if !procedureKinds[sym.Kind] || sym.Name == "" {
				continue
			}
			candidate := Candidate{
				QualifiedName: sym.Module + "." + sym.Name,
				Kind:          sym.Kind,
				File:          sym.File,
				Line:          sym.StartLine,
			}
			key := strings.ToLower(sym.Name)
			res.byName[key] = append(res.byName[key], candidate)
		}
	}
	for key := range res.byName {
		sort.Slice(res.byName[key], func(i, j int) bool {
			return res.byName[key][i].QualifiedName < res.byName[key][j].QualifiedName
		})
	}
	return res
}

func (e *extractor) visit(node *tree_sitter.Node) {
	if node == nil {
		return
	}
	switch node.Kind() {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration":
		prior := e.current
		e.current = e.callerForProcedure(node)
		for i := uint(0); i < node.NamedChildCount(); i++ {
			e.visit(node.NamedChild(i))
		}
		e.current = prior
		return
	case "call_statement":
		e.addCallFromNode(node, "callee")
		e.visitCallStatementChildren(node)
		return
	case "call_expression":
		e.addCallFromNode(node, "function")
		e.visitCallExpressionArguments(node)
		return
	case "new_expression":
		e.addNewExpression(node)
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		e.visit(node.NamedChild(i))
	}
}

func (e *extractor) visitCallStatementChildren(node *tree_sitter.Node) {
	callee := node.ChildByFieldName("callee")
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		if sameNode(child, callee) {
			if child.Kind() == "call_expression" {
				e.visitCallExpressionArguments(child)
			}
			continue
		}
		e.visit(child)
	}
}

func (e *extractor) visitCallExpressionArguments(node *tree_sitter.Node) {
	fn := node.ChildByFieldName("function")
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || sameNode(child, fn) {
			continue
		}
		e.visit(child)
	}
}

func (e *extractor) addCallFromNode(node *tree_sitter.Node, field string) {
	target := node.ChildByFieldName(field)
	if target == nil && node.Kind() == "call_statement" {
		target = firstNamedChild(node)
	}
	if target == nil {
		return
	}
	callNode := node
	argumentSource := target
	if target.Kind() == "call_expression" {
		if fn := target.ChildByFieldName("function"); fn != nil {
			target = fn
		}
	}
	callee := calleeFromNode(target, e.source)
	if callee.Text == "" {
		return
	}
	args := argumentsFromCallNode(callNode, argumentSource, e.source)
	resolution := e.resolver.resolve(callee)
	e.calls = append(e.calls, Call{
		File:       e.file,
		Module:     e.moduleName,
		Caller:     cloneCaller(e.current),
		Callee:     callee,
		Arguments:  args,
		Range:      vbaast.NodeRange(callNode),
		Parse:      e.parse,
		Resolution: resolution,
	})
}

func (e *extractor) addNewExpression(node *tree_sitter.Node) {
	target := node.ChildByFieldName("type")
	if target == nil {
		return
	}
	typ := strings.TrimSpace(target.Utf8Text(e.source))
	if typ == "" {
		return
	}
	callee := calleeFromNode(target, e.source)
	if callee.Text == "" {
		callee.Text = typ
		callee.BaseName = lastNamePart(typ)
		callee.Member = callee.BaseName
	}
	callee.Text = "New " + callee.Text
	e.calls = append(e.calls, Call{
		File:       e.file,
		Module:     e.moduleName,
		Caller:     cloneCaller(e.current),
		Callee:     callee,
		Arguments:  Arguments{Named: []NamedArgument{}},
		Range:      vbaast.NodeRange(node),
		Parse:      e.parse,
		Resolution: e.resolver.resolve(callee),
	})
}

func (e *extractor) callerForProcedure(node *tree_sitter.Node) *Caller {
	name := ""
	if nameNode := node.ChildByFieldName("name"); nameNode != nil {
		name = nameNode.Utf8Text(e.source)
	}
	kind := strings.TrimSuffix(node.Kind(), "_declaration")
	switch node.Kind() {
	case "property_get_declaration":
		kind = "property_get"
	case "property_let_declaration":
		kind = "property_let"
	case "property_set_declaration":
		kind = "property_set"
	}
	qualified := name
	if e.moduleName != "" && name != "" {
		qualified = e.moduleName + "." + name
	}
	return &Caller{Name: name, Kind: kind, QualifiedName: qualified}
}

func calleeFromNode(node *tree_sitter.Node, source []byte) Callee {
	text := strings.TrimSpace(node.Utf8Text(source))
	callee := Callee{Text: text}
	switch node.Kind() {
	case "qualified_member_expression":
		object := node.ChildByFieldName("object")
		property := node.ChildByFieldName("property")
		if object != nil {
			receiver := strings.TrimSpace(object.Utf8Text(source))
			callee.Receiver = &receiver
		}
		if property != nil {
			callee.Member = strings.TrimSpace(property.Utf8Text(source))
			callee.BaseName = cleanIdentifier(callee.Member)
		}
	case "implicit_member_expression":
		property := node.ChildByFieldName("property")
		if property != nil {
			callee.Member = strings.TrimSpace(property.Utf8Text(source))
			callee.BaseName = cleanIdentifier(callee.Member)
		}
	default:
		callee.BaseName = cleanIdentifier(lastNamePart(text))
		callee.Member = callee.BaseName
	}
	if callee.BaseName == "" {
		callee.BaseName = cleanIdentifier(lastNamePart(text))
	}
	if callee.Member == "" {
		callee.Member = callee.BaseName
	}
	return callee
}

func argumentsFromCallNode(callNode, target *tree_sitter.Node, source []byte) Arguments {
	args := Arguments{Named: []NamedArgument{}}
	if callNode.Kind() == "call_expression" {
		return argumentsFromCallExpression(callNode, source)
	}
	if target != nil && target.Kind() == "call_expression" {
		return argumentsFromCallExpression(target, source)
	}
	for i := uint(0); i < callNode.NamedChildCount(); i++ {
		child := callNode.NamedChild(i)
		if child == nil || sameNode(child, target) {
			continue
		}
		if child.Kind() == "argument_list" {
			return argumentsFromArgumentList(child, source)
		}
	}
	return args
}

func argumentsFromCallExpression(node *tree_sitter.Node, source []byte) Arguments {
	args := Arguments{Named: []NamedArgument{}}
	fn := node.ChildByFieldName("function")
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || sameNode(child, fn) {
			continue
		}
		args.Count++
		if child.Kind() == "named_argument" {
			args.Named = append(args.Named, namedArgument(child, source))
		}
	}
	return args
}

func argumentsFromArgumentList(node *tree_sitter.Node, source []byte) Arguments {
	args := Arguments{Named: []NamedArgument{}}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		args.Count++
		if child.Kind() == "named_argument" {
			args.Named = append(args.Named, namedArgument(child, source))
		}
	}
	return args
}

func namedArgument(node *tree_sitter.Node, source []byte) NamedArgument {
	name := ""
	if nameNode := node.ChildByFieldName("name"); nameNode != nil {
		name = cleanIdentifier(nameNode.Utf8Text(source))
	}
	value := ""
	if valueNode := node.ChildByFieldName("value"); valueNode != nil {
		value = strings.TrimSpace(valueNode.Utf8Text(source))
	}
	return NamedArgument{Name: name, ValueText: value}
}

func (r resolver) resolve(callee Callee) Resolution {
	base := strings.TrimPrefix(callee.BaseName, "New ")
	base = cleanIdentifier(base)
	candidates := r.byName[strings.ToLower(base)]
	if callee.Receiver != nil {
		receiver := cleanQualifiedName(*callee.Receiver)
		if isExternalLikeReceiver(receiver) {
			return Resolution{Status: "external"}
		}
		matches := candidatesForReceiver(candidates, receiver, base)
		if len(matches) == 1 {
			return Resolution{Status: "matched", Candidates: matches}
		}
		if len(matches) > 1 {
			return Resolution{Status: "ambiguous", Candidates: matches}
		}
		return Resolution{Status: "member_call"}
	}
	if base != "" {
		if len(candidates) == 1 {
			return Resolution{Status: "matched", Candidates: candidates}
		}
		if len(candidates) > 1 {
			return Resolution{Status: "ambiguous", Candidates: candidates}
		}
	}
	textKey := strings.ToLower(strings.TrimPrefix(callee.Text, "New "))
	if builtinLikeNames[textKey] || builtinLikeNames[strings.ToLower(base)] {
		return Resolution{Status: "builtin_like"}
	}
	return Resolution{Status: "unresolved"}
}

func candidatesForReceiver(candidates []Candidate, receiver, base string) []Candidate {
	if receiver == "" || base == "" {
		return nil
	}
	qualified := strings.ToLower(receiver + "." + base)
	shortQualified := strings.ToLower(cleanIdentifier(lastNamePart(receiver)) + "." + base)
	matches := make([]Candidate, 0, len(candidates))
	for _, candidate := range candidates {
		name := strings.ToLower(candidate.QualifiedName)
		if name == qualified || name == shortQualified {
			matches = append(matches, candidate)
		}
	}
	return matches
}

func isExternalLikeReceiver(receiver string) bool {
	for _, part := range strings.FieldsFunc(receiver, func(r rune) bool {
		return r == '.' || r == '!'
	}) {
		if externalLikeReceivers[strings.ToLower(cleanIdentifier(part))] {
			return true
		}
	}
	return false
}

func addResolutionSummary(summary *ResultSummary, status string) {
	switch status {
	case "matched":
		summary.Matched++
	case "ambiguous":
		summary.Ambiguous++
	case "external":
		summary.External++
	case "builtin_like":
		summary.BuiltinLike++
	case "member_call":
		summary.MemberCalls++
	default:
		summary.Unresolved++
	}
}

func matchesFrom(call Call, filter string) bool {
	filter = strings.TrimSpace(filter)
	if filter == "" {
		return true
	}
	if call.Caller == nil {
		return false
	}
	return strings.EqualFold(call.Caller.Name, filter) ||
		strings.EqualFold(call.Caller.QualifiedName, filter) ||
		strings.EqualFold(call.Module, filter)
}

func matchesTo(call Call, filter string) bool {
	filter = strings.TrimSpace(filter)
	if filter == "" {
		return true
	}
	return strings.EqualFold(call.Callee.BaseName, filter) ||
		strings.EqualFold(call.Callee.Member, filter) ||
		strings.EqualFold(call.Callee.Text, filter)
}

func resolveDisplayPath(root, path string) string {
	if filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(root, filepath.FromSlash(path))
}

func firstNamedChild(node *tree_sitter.Node) *tree_sitter.Node {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		if child := node.NamedChild(i); child != nil {
			return child
		}
	}
	return nil
}

func sameNode(a, b *tree_sitter.Node) bool {
	if a == nil || b == nil {
		return false
	}
	return a.StartByte() == b.StartByte() && a.EndByte() == b.EndByte() && a.Kind() == b.Kind()
}

func cloneCaller(caller *Caller) *Caller {
	if caller == nil {
		return nil
	}
	clone := *caller
	return &clone
}

func lastNamePart(text string) string {
	text = strings.TrimSpace(text)
	text = strings.TrimPrefix(text, "New ")
	for _, sep := range []string{".", "!"} {
		if idx := strings.LastIndex(text, sep); idx >= 0 {
			text = text[idx+1:]
		}
	}
	return text
}

func cleanIdentifier(text string) string {
	text = strings.TrimSpace(text)
	text = strings.Trim(text, "[]")
	text = strings.TrimRight(text, "$%&#@^!")
	return text
}

func cleanQualifiedName(text string) string {
	parts := strings.FieldsFunc(strings.TrimSpace(text), func(r rune) bool {
		return r == '.' || r == '!'
	})
	cleaned := make([]string, 0, len(parts))
	for _, part := range parts {
		if clean := cleanIdentifier(part); clean != "" {
			cleaned = append(cleaned, clean)
		}
	}
	return strings.Join(cleaned, ".")
}

func (c Call) String() string {
	caller := "<module>"
	if c.Caller != nil && c.Caller.QualifiedName != "" {
		caller = c.Caller.QualifiedName
	}
	return fmt.Sprintf("%s -> %s %s:%d", caller, c.Callee.Text, c.File, c.Range.StartLine)
}
