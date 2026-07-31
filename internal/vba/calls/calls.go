package calls

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
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
	// Symbols is the complete, private-inclusive project symbol snapshot used to
	// resolve Calls. It is intentionally not part of inspect calls JSON, but lets
	// higher-level analyses reuse this parsed project snapshot without a second
	// source walk.
	Symbols     []symbols.Symbol  `json:"-"`
	ModuleKinds map[string]string `json:"-"`
	// TypeReferences are syntax-local project-type facts retained for higher
	// level projections without changing inspect calls JSON.
	TypeReferences []TypeReference `json:"-"`
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

// SourceOptions configures syntax-local call-site extraction from one VBA
// source document.
type SourceOptions struct {
	RootDir    string
	Path       string
	ModuleKind string
}

// FileResult contains syntax-local call sites extracted from one VBA source
// document. It remains useful for files with no calls because Parse records the
// document's recovery state.
type FileResult struct {
	Path           string               `json:"path"`
	ModuleName     string               `json:"moduleName"`
	ModuleKind     string               `json:"moduleKind"`
	Parse          symbols.ParseSummary `json:"parse"`
	CallSites      []CallSite           `json:"callSites"`
	TypeReferences []TypeReference      `json:"-"`
}

// CallSite contains only facts available from one parsed VBA document. It does
// not depend on project symbols and never retains tree-sitter nodes.
type CallSite struct {
	File      string               `json:"file"`
	Module    string               `json:"module"`
	Caller    *Caller              `json:"caller,omitempty"`
	Callee    Callee               `json:"callee"`
	Arguments Arguments            `json:"arguments"`
	Range     vbaast.Range         `json:"range"`
	Parse     symbols.ParseSummary `json:"parse"`
}

// TypeReference records an explicit VBA type context from one parsed document.
// Kind is uses_type, constructs, or implements.
type TypeReference struct {
	Kind   string               `json:"kind"`
	File   string               `json:"file"`
	Module string               `json:"module"`
	Caller *Caller              `json:"caller,omitempty"`
	Target string               `json:"target"`
	Range  vbaast.Range         `json:"range"`
	Parse  symbols.ParseSummary `json:"parse"`
}

// Call is a syntax-local CallSite resolved against project symbols. Embedding
// keeps the existing flat JSON representation backward compatible.
type Call struct {
	CallSite
	Resolution Resolution `json:"resolution"`
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

	module     string
	visibility string
}

// ResolverSymbol is the protocol-neutral symbol shape needed to resolve call
// sites. Line is the 1-based declaration line reported in resolved candidates.
type ResolverSymbol struct {
	Name       string
	Module     string
	ModuleKind string
	Kind       string
	Visibility string
	File       string
	Line       int
}

type extractor struct {
	source     []byte
	file       string
	moduleName string
	parse      symbols.ParseSummary
	current    *Caller
	callSites  []CallSite
	typeRefs   []TypeReference
}

// Resolver resolves raw call sites against a snapshot of project procedure
// symbols. A Resolver can be replaced without re-extracting unchanged sites.
type Resolver struct {
	byName  map[string][]Candidate
	symbols procedureir.SymbolResolver
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
	"debug.print": true, "dir": true, "doevents": true, "environ": true, "error": true,
	"format": true, "getobject": true, "inputbox": true, "instr": true,
	"isarray": true, "isdate": true, "isempty": true, "iserror": true,
	"isnull": true, "isnumeric": true, "join": true, "lbound": true,
	"lcase": true, "left": true, "len": true, "mid": true, "msgbox": true,
	"replace": true, "right": true, "rnd": true, "shell": true, "split": true, "str": true,
	"trim": true, "typename": true, "ubound": true, "ucase": true, "val": true,
}

var externalLikeReceivers = map[string]bool{
	"application": true, "debug": true, "excel": true, "worksheetfunction": true,
}

var moduleNameAttributeRe = regexp.MustCompile(`(?i)^\s*Attribute\s+VB_Name\s*=\s*(.*)\s*$`)

func Inspect(opts Options) (*Result, error) {
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	absRoot, err := filepath.Abs(rootDir)
	if err != nil {
		return nil, err
	}

	files, err := symbols.DiscoverSourceFiles(symbols.Options{
		RootDir:        absRoot,
		Config:         opts.Config,
		Path:           opts.Path,
		IncludePrivate: true,
		IncludeLabels:  false,
	})
	if err != nil {
		return nil, err
	}

	displayRoot := opts.Path
	if strings.TrimSpace(displayRoot) == "" {
		displayRoot = "src"
	}
	result := &Result{Root: filepath.ToSlash(displayRoot), Calls: []Call{}, Symbols: []symbols.Symbol{}, ModuleKinds: map[string]string{}}
	allSymbols := make([]symbols.Symbol, 0)
	allSites := make([]CallSite, 0)
	for _, file := range files {
		source, err := os.ReadFile(file.Path)
		if err != nil {
			return nil, err
		}
		doc, err := vbaast.ParseDocument(file.Path, source)
		if err != nil {
			return nil, err
		}
		symbolFile, err := symbols.InspectParsed(symbols.SourceOptions{
			RootDir:        absRoot,
			Path:           file.Path,
			ModuleKind:     file.ModuleKind,
			IncludePrivate: true,
			IncludeLabels:  false,
		}, doc)
		if err != nil {
			doc.Close()
			return nil, err
		}
		callFile, err := ExtractParsed(SourceOptions{
			RootDir:    absRoot,
			Path:       file.Path,
			ModuleKind: file.ModuleKind,
		}, doc)
		doc.Close()
		if err != nil {
			return nil, err
		}
		allSymbols = append(allSymbols, symbolFile.Symbols...)
		for _, sym := range symbolFile.Symbols {
			if sym.Kind == "module" {
				result.ModuleKinds[sym.File] = file.ModuleKind
			}
		}
		allSites = append(allSites, callFile.CallSites...)
		result.TypeReferences = append(result.TypeReferences, callFile.TypeReferences...)
		if callFile.Parse.HasError {
			result.Summary.ParseErrors++
		}
		if callFile.Parse.HasMissing {
			result.Summary.MissingNodes++
		}
	}
	resolver := NewResolver(allSymbols)
	result.Symbols = append(result.Symbols, allSymbols...)
	for _, site := range allSites {
		call := resolver.Resolve(site)
		if !matchesFrom(call, opts.From) || !matchesTo(call, opts.To) {
			continue
		}
		result.Calls = append(result.Calls, call)
		addResolutionSummary(&result.Summary, call.Resolution.Status)
	}
	result.Summary.Files = len(files)
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

// ExtractSource parses source, extracts raw call sites, and closes the parsed
// document before returning.
func ExtractSource(opts SourceOptions, source []byte) (FileResult, error) {
	path := opts.Path
	if strings.TrimSpace(path) == "" {
		path = "Untitled.bas"
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return FileResult{}, err
	}
	defer doc.Close()
	return ExtractParsed(opts, doc)
}

// ExtractParsed extracts raw call sites from a caller-owned parsed VBA
// document. It does not close doc or retain tree-sitter nodes after Read
// returns.
func ExtractParsed(opts SourceOptions, doc *vbaast.ParsedDocument) (FileResult, error) {
	rootDir := opts.RootDir
	if strings.TrimSpace(rootDir) == "" {
		rootDir = "."
	}
	rootDir, err := filepath.Abs(rootDir)
	if err != nil {
		return FileResult{}, err
	}
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{
		RootDir:    rootDir,
		Path:       opts.Path,
		ModuleKind: opts.ModuleKind,
	}, doc)
	if err != nil {
		return FileResult{}, err
	}
	return ExtractIR(ir), nil
}

// ExtractIR projects syntax-local calls and type references from the shared
// procedure IR without reparsing or rescanning source.
func ExtractIR(ir procedureir.DocumentIR) FileResult {
	result := FileResult{
		Path:       ir.Path,
		ModuleName: ir.ModuleName,
		ModuleKind: ir.ModuleKind,
		Parse: symbols.ParseSummary{
			HasError: ir.Parse.HasError, HasMissing: ir.Parse.HasMissing,
		},
		CallSites:      []CallSite{},
		TypeReferences: []TypeReference{},
	}
	for _, procedure := range ir.Procedures {
		for _, site := range procedure.Calls {
			result.CallSites = append(result.CallSites, callSiteFromIR(site, result.Parse))
		}
	}
	for _, reference := range ir.TypeReferences {
		result.TypeReferences = append(result.TypeReferences, typeReferenceFromIR(reference, result.Parse))
	}
	sort.SliceStable(result.CallSites, func(i, j int) bool {
		return result.CallSites[i].Range.StartByte < result.CallSites[j].Range.StartByte
	})
	sort.SliceStable(result.TypeReferences, func(i, j int) bool {
		return result.TypeReferences[i].Range.StartByte < result.TypeReferences[j].Range.StartByte
	})
	return result
}

func callSiteFromIR(site procedureir.CallSite, parse symbols.ParseSummary) CallSite {
	var caller *Caller
	if site.Caller.Name != "" || site.Caller.QualifiedName != "" {
		caller = &Caller{
			Name: site.Caller.Name, Kind: string(site.Caller.Kind),
			QualifiedName: site.Caller.QualifiedName,
		}
	}
	named := make([]NamedArgument, len(site.Arguments.Named))
	for i, argument := range site.Arguments.Named {
		named[i] = NamedArgument{Name: argument.Name, ValueText: argument.ValueText}
	}
	return CallSite{
		File: site.File, Module: site.Module, Caller: caller,
		Callee: Callee{
			Text: site.Callee.Text, BaseName: site.Callee.BaseName,
			Receiver: cloneStringPointer(site.Callee.Receiver), Member: site.Callee.Member,
		},
		Arguments: Arguments{Count: site.Arguments.Count, Named: named},
		Range:     site.Range, Parse: parse,
	}
}

func typeReferenceFromIR(reference procedureir.TypeReference, parse symbols.ParseSummary) TypeReference {
	var caller *Caller
	if reference.Caller != nil {
		caller = &Caller{
			Name: reference.Caller.Name, Kind: string(reference.Caller.Kind),
			QualifiedName: reference.Caller.QualifiedName,
		}
	}
	return TypeReference{
		Kind: reference.Kind, File: reference.File, Module: reference.Module,
		Caller: caller, Target: reference.Target, Range: reference.Range, Parse: parse,
	}
}

func cloneStringPointer(value *string) *string {
	if value == nil {
		return nil
	}
	clone := *value
	return &clone
}

// extractParsedLegacy is retained temporarily as an independent compatibility
// oracle for tests while call extraction migrates to procedureir.
func extractParsedLegacy(opts SourceOptions, doc *vbaast.ParsedDocument) (FileResult, error) {
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	rootDir, err := filepath.Abs(rootDir)
	if err != nil {
		return FileResult{}, err
	}
	var result FileResult
	err = doc.Read(func(view vbaast.ParsedView) error {
		path := opts.Path
		if strings.TrimSpace(path) == "" {
			path = view.Path
		}
		if strings.TrimSpace(path) == "" {
			path = "Untitled.bas"
		}
		rel := displayPath(rootDir, path)
		if !filepath.IsAbs(path) {
			rel = filepath.ToSlash(path)
		}
		moduleName := moduleNameFromSource(path, view.Source)
		moduleKind := opts.ModuleKind
		if moduleKind == "" {
			moduleKind = moduleKindFromPath(path)
		}
		parse := symbols.ParseSummary{
			HasError:   view.HasError,
			HasMissing: view.HasMissing,
		}
		ext := extractor{
			source:     view.Source,
			file:       rel,
			moduleName: moduleName,
			parse:      parse,
		}
		ext.visit(view.Root)
		result = FileResult{
			Path:           rel,
			ModuleName:     moduleName,
			ModuleKind:     moduleKind,
			Parse:          parse,
			CallSites:      ext.callSites,
			TypeReferences: ext.typeRefs,
		}
		if result.CallSites == nil {
			result.CallSites = []CallSite{}
		}
		if result.TypeReferences == nil {
			result.TypeReferences = []TypeReference{}
		}
		return nil
	})
	return result, err
}

// NewResolver creates a deterministic project-symbol resolver.
func NewResolver(projectSymbols []symbols.Symbol) Resolver {
	resolverSymbols := make([]ResolverSymbol, 0, len(projectSymbols))
	for _, sym := range projectSymbols {
		resolverSymbols = append(resolverSymbols, ResolverSymbol{
			Name:       sym.Name,
			Module:     sym.Module,
			ModuleKind: sym.ModuleKind,
			Kind:       sym.Kind,
			Visibility: sym.Visibility,
			File:       sym.File,
			Line:       sym.StartLine,
		})
	}
	return NewResolverFromSymbols(resolverSymbols)
}

// NewResolverFromSymbols creates a deterministic resolver from the
// protocol-neutral symbol shape.
func NewResolverFromSymbols(projectSymbols []ResolverSymbol) Resolver {
	irSymbols := make([]procedureir.ResolverSymbol, 0, len(projectSymbols))
	for _, sym := range projectSymbols {
		irSymbols = append(irSymbols, procedureir.ResolverSymbol{
			Name:       sym.Name,
			Module:     sym.Module,
			ModuleKind: sym.ModuleKind,
			Kind:       sym.Kind,
			Visibility: sym.Visibility,
			File:       sym.File,
			Line:       sym.Line,
		})
	}
	res := Resolver{
		byName:  map[string][]Candidate{},
		symbols: procedureir.NewResolver(irSymbols),
	}
	for _, sym := range projectSymbols {
		if !procedureKinds[sym.Kind] || sym.Name == "" {
			continue
		}
		candidate := Candidate{
			QualifiedName: sym.Module + "." + sym.Name,
			Kind:          sym.Kind,
			File:          normalizeCandidateFile(sym.File),
			Line:          sym.Line,
		}
		if strings.EqualFold(sym.Visibility, "private") {
			candidate.module = sym.Module
			candidate.visibility = sym.Visibility
		}
		key := strings.ToLower(sym.Name)
		res.byName[key] = append(res.byName[key], candidate)
	}
	for key := range res.byName {
		sort.Slice(res.byName[key], func(i, j int) bool {
			return candidateLess(res.byName[key][i], res.byName[key][j])
		})
	}
	return res
}

// Resolve attaches the resolution derived from this Resolver without mutating
// the raw call site.
func (r Resolver) Resolve(site CallSite) Call {
	return Call{
		CallSite:   CloneCallSite(site),
		Resolution: r.resolveCallee(site),
	}
}

// ResolveCall adapts the existing project call resolver to procedureir without
// coupling the IR package to inspect-call result types.
func (r Resolver) ResolveCall(site procedureir.CallSite) procedureir.CallResolution {
	resolved := r.Resolve(callSiteFromIR(site, symbols.ParseSummary{})).Resolution
	candidates := make([]procedureir.Candidate, len(resolved.Candidates))
	for i, candidate := range resolved.Candidates {
		candidates[i] = procedureir.Candidate{
			QualifiedName: candidate.QualifiedName,
			Kind:          candidate.Kind,
			File:          candidate.File,
			Line:          candidate.Line,
		}
	}
	return procedureir.CallResolution{
		Status:     procedureir.ResolutionStatus(resolved.Status),
		Candidates: candidates,
	}
}

// ResolveSymbol provides the symbol half of procedureir.Resolver using the
// complete project symbol snapshot, including non-procedure declarations.
func (r Resolver) ResolveSymbol(reference procedureir.SymbolReference) procedureir.SymbolResolution {
	return r.symbols.ResolveSymbol(reference)
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
		e.addTypeReference("constructs", node, "type")
	case "as_type_clause":
		e.addTypeReference("uses_type", node, "type")
		if containsKeyword(node.Utf8Text(e.source), "New") {
			e.addTypeReference("constructs", node, "type")
		}
	case "implements_statement":
		e.addTypeReference("implements", node, "name")
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		e.visit(node.NamedChild(i))
	}
}

func (e *extractor) addTypeReference(kind string, node *tree_sitter.Node, field string) {
	target := node.ChildByFieldName(field)
	if target == nil {
		return
	}
	text := strings.TrimSpace(target.Utf8Text(e.source))
	if text == "" {
		return
	}
	e.typeRefs = append(e.typeRefs, TypeReference{
		Kind: kind, File: e.file, Module: e.moduleName, Caller: cloneCaller(e.current),
		Target: text, Range: vbaast.NodeRange(target), Parse: e.parse,
	})
}

func containsKeyword(text, keyword string) bool {
	for _, part := range strings.Fields(text) {
		if strings.EqualFold(part, keyword) {
			return true
		}
	}
	return false
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
	e.callSites = append(e.callSites, CallSite{
		File:      e.file,
		Module:    e.moduleName,
		Caller:    cloneCaller(e.current),
		Callee:    callee,
		Arguments: args,
		Range:     vbaast.NodeRange(callNode),
		Parse:     e.parse,
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
	e.callSites = append(e.callSites, CallSite{
		File:      e.file,
		Module:    e.moduleName,
		Caller:    cloneCaller(e.current),
		Callee:    callee,
		Arguments: Arguments{Named: []NamedArgument{}},
		Range:     vbaast.NodeRange(node),
		Parse:     e.parse,
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
		receiverNode := childByFieldNameAny(node, "receiver", "object")
		memberNode := childByFieldNameAny(node, "member", "property")
		if receiverNode != nil {
			receiver := strings.TrimSpace(receiverNode.Utf8Text(source))
			callee.Receiver = &receiver
		}
		if memberNode != nil {
			callee.Member = strings.TrimSpace(memberNode.Utf8Text(source))
			callee.BaseName = cleanIdentifier(callee.Member)
		}
	case "implicit_member_expression":
		memberNode := childByFieldNameAny(node, "member", "property")
		if memberNode != nil {
			callee.Member = strings.TrimSpace(memberNode.Utf8Text(source))
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
	if list := callNode.ChildByFieldName("arguments"); list != nil {
		return argumentsFromArgumentList(list, source)
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
	if list := node.ChildByFieldName("arguments"); list != nil {
		return argumentsFromArgumentList(list, source)
	}
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

func childByFieldNameAny(node *tree_sitter.Node, names ...string) *tree_sitter.Node {
	if node == nil {
		return nil
	}
	for _, name := range names {
		if child := node.ChildByFieldName(name); child != nil {
			return child
		}
	}
	return nil
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

func (r Resolver) resolveCallee(site CallSite) Resolution {
	callee := site.Callee
	base := strings.TrimPrefix(callee.BaseName, "New ")
	base = cleanIdentifier(base)
	candidates := visibleCandidates(r.byName[strings.ToLower(base)], site.Caller)
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
			return Resolution{Status: "matched", Candidates: cloneCandidates(candidates)}
		}
		if len(candidates) > 1 {
			return Resolution{Status: "ambiguous", Candidates: cloneCandidates(candidates)}
		}
	}
	textKey := strings.ToLower(strings.TrimPrefix(callee.Text, "New "))
	if builtinLikeNames[textKey] || builtinLikeNames[strings.ToLower(base)] {
		return Resolution{Status: "builtin_like"}
	}
	return Resolution{Status: "unresolved"}
}

// visibleCandidates excludes private procedures outside the caller's module.
// A module-level call has no procedure caller, so it cannot target a private
// procedure either.
func visibleCandidates(candidates []Candidate, caller *Caller) []Candidate {
	if len(candidates) == 0 {
		return nil
	}
	visible := make([]Candidate, 0, len(candidates))
	for _, candidate := range candidates {
		if strings.EqualFold(candidate.visibility, "private") &&
			(caller == nil || !strings.EqualFold(candidate.module, callerModule(caller))) {
			continue
		}
		visible = append(visible, candidate)
	}
	return visible
}

func callerModule(caller *Caller) string {
	if caller == nil {
		return ""
	}
	qualified := strings.TrimSpace(caller.QualifiedName)
	if index := strings.LastIndex(qualified, "."); index > 0 {
		return qualified[:index]
	}
	return ""
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

func moduleNameFromSource(path string, source []byte) string {
	for _, line := range strings.Split(string(source), "\n") {
		match := moduleNameAttributeRe.FindStringSubmatch(line)
		if match == nil {
			continue
		}
		value := strings.Trim(strings.TrimSpace(match[1]), `"`)
		if value != "" {
			return value
		}
	}
	return strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
}

func moduleKindFromPath(path string) string {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".cls":
		return "class"
	case ".frm":
		return "form"
	default:
		return "standard"
	}
}

func cloneCandidates(candidates []Candidate) []Candidate {
	if len(candidates) == 0 {
		return nil
	}
	return append([]Candidate(nil), candidates...)
}

func displayPath(rootDir, path string) string {
	rel, err := filepath.Rel(rootDir, path)
	if err != nil {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
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

// CloneCallSite returns a deep copy suitable for use across cache and resolver
// boundaries.
func CloneCallSite(site CallSite) CallSite {
	clone := site
	clone.Caller = cloneCaller(site.Caller)
	if site.Callee.Receiver != nil {
		receiver := *site.Callee.Receiver
		clone.Callee.Receiver = &receiver
	}
	if site.Arguments.Named != nil {
		clone.Arguments.Named = make([]NamedArgument, len(site.Arguments.Named))
		copy(clone.Arguments.Named, site.Arguments.Named)
	}
	return clone
}

// CloneFileResult returns a deep copy suitable for storing in or returning
// from a cache without exposing mutable call-site state.
func CloneFileResult(result FileResult) FileResult {
	clone := result
	if result.CallSites != nil {
		clone.CallSites = make([]CallSite, len(result.CallSites))
		for i := range result.CallSites {
			clone.CallSites[i] = CloneCallSite(result.CallSites[i])
		}
	}
	if result.TypeReferences != nil {
		clone.TypeReferences = make([]TypeReference, len(result.TypeReferences))
		for i := range result.TypeReferences {
			clone.TypeReferences[i] = result.TypeReferences[i]
			clone.TypeReferences[i].Caller = cloneCaller(result.TypeReferences[i].Caller)
		}
	}
	return clone
}

func normalizeCandidateFile(file string) string {
	if strings.TrimSpace(file) == "" {
		return ""
	}
	return filepath.ToSlash(filepath.Clean(file))
}

func candidateLess(a, b Candidate) bool {
	if a.QualifiedName != b.QualifiedName {
		return a.QualifiedName < b.QualifiedName
	}
	if a.Kind != b.Kind {
		return a.Kind < b.Kind
	}
	if a.File != b.File {
		return a.File < b.File
	}
	return a.Line < b.Line
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
