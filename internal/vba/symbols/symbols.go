package symbols

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Options struct {
	RootDir        string
	Config         config.Config
	Path           string
	IncludePrivate bool
	IncludeLabels  bool
	Module         string
}

type Result struct {
	Root    string        `json:"root"`
	Files   []FileResult  `json:"files"`
	Summary ResultSummary `json:"summary"`
}

type ResultSummary struct {
	Files        int `json:"files"`
	Symbols      int `json:"symbols"`
	ParseErrors  int `json:"parseErrors"`
	MissingNodes int `json:"missingNodes"`
}

type FileResult struct {
	Path       string       `json:"path"`
	ModuleName string       `json:"moduleName"`
	ModuleKind string       `json:"moduleKind"`
	Parse      ParseSummary `json:"parse"`
	Symbols    []Symbol     `json:"symbols"`
}

type ParseSummary struct {
	HasError   bool `json:"hasError"`
	HasMissing bool `json:"hasMissing"`
}

type Symbol struct {
	Name        string      `json:"name"`
	Kind        string      `json:"kind"`
	Visibility  string      `json:"visibility,omitempty"`
	Module      string      `json:"module"`
	File        string      `json:"file"`
	Parent      string      `json:"parent,omitempty"`
	StartLine   int         `json:"startLine"`
	StartColumn int         `json:"startColumn"`
	EndLine     int         `json:"endLine"`
	EndColumn   int         `json:"endColumn"`
	StartByte   int         `json:"startByte"`
	EndByte     int         `json:"endByte"`
	Signature   string      `json:"signature,omitempty"`
	Attributes  []Attribute `json:"attributes,omitempty"`
	Static      bool        `json:"static,omitempty"`
	ReturnType  string      `json:"returnType,omitempty"`
	Parameters  []Parameter `json:"parameters,omitempty"`
}

type Attribute struct {
	Name  string `json:"name"`
	Value string `json:"value,omitempty"`
}

type Parameter struct {
	Name     string  `json:"name"`
	Type     string  `json:"type,omitempty"`
	Passing  string  `json:"passing,omitempty"`
	Optional bool    `json:"optional,omitempty"`
	Default  *string `json:"default"`
}

type fileCandidate struct {
	path       string
	moduleKind string
}

type SourceOptions struct {
	RootDir        string
	Path           string
	ModuleKind     string
	IncludePrivate bool
	IncludeLabels  bool
}

type extractor struct {
	opts       Options
	rootDir    string
	source     []byte
	file       string
	moduleName string
	moduleKind string
	attrs      []Attribute
	symbols    []Symbol
}

var attrRe = regexp.MustCompile(`(?i)^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$`)

func Inspect(opts Options) (*Result, error) {
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	rootDir, err := filepath.Abs(rootDir)
	if err != nil {
		return nil, err
	}
	opts.RootDir = rootDir
	displayRoot := opts.Path
	if strings.TrimSpace(displayRoot) == "" {
		displayRoot = "src"
	}

	files, err := discoverFiles(opts)
	if err != nil {
		return nil, err
	}
	parser, err := vbaast.NewParser()
	if err != nil {
		return nil, err
	}
	defer parser.Close()

	result := &Result{
		Root:  filepath.ToSlash(displayRoot),
		Files: []FileResult{},
	}
	moduleFilter := strings.TrimSpace(opts.Module)
	for _, file := range files {
		parsed, err := parser.ParseFile(file.path)
		if err != nil {
			return nil, err
		}
		rel := displayPath(rootDir, file.path)
		moduleName, attrs := moduleMetadata(file.path, parsed.Source)
		if moduleFilter != "" && !strings.EqualFold(moduleFilter, moduleName) {
			parsed.Close()
			continue
		}
		ext := extractor{
			opts:       opts,
			rootDir:    rootDir,
			source:     parsed.Source,
			file:       rel,
			moduleName: moduleName,
			moduleKind: file.moduleKind,
			attrs:      attrs,
		}
		fileSymbols := ext.extract(parsed.Root)
		result.Files = append(result.Files, FileResult{
			Path:       rel,
			ModuleName: moduleName,
			ModuleKind: file.moduleKind,
			Parse: ParseSummary{
				HasError:   parsed.HasError,
				HasMissing: parsed.HasMissing,
			},
			Symbols: fileSymbols,
		})
		result.Summary.Symbols += len(fileSymbols)
		if parsed.HasError {
			result.Summary.ParseErrors++
		}
		if parsed.HasMissing {
			result.Summary.MissingNodes++
		}
		parsed.Close()
	}
	result.Summary.Files = len(result.Files)
	return result, nil
}

func InspectSource(opts SourceOptions, source []byte) (FileResult, error) {
	rootDir := opts.RootDir
	if rootDir == "" {
		rootDir = "."
	}
	rootDir, err := filepath.Abs(rootDir)
	if err != nil {
		return FileResult{}, err
	}
	path := opts.Path
	if strings.TrimSpace(path) == "" {
		path = "Untitled.bas"
	}
	moduleKind := opts.ModuleKind
	if moduleKind == "" {
		moduleKind = kindForPath(rootDir, config.Config{}, path)
	}
	parser, err := vbaast.NewParser()
	if err != nil {
		return FileResult{}, err
	}
	defer parser.Close()
	parsed := parser.Parse(path, source)
	defer parsed.Close()
	rel := displayPath(rootDir, path)
	if !filepath.IsAbs(path) {
		rel = filepath.ToSlash(path)
	}
	moduleName, attrs := moduleMetadata(path, parsed.Source)
	ext := extractor{
		opts: Options{
			RootDir:        rootDir,
			IncludePrivate: opts.IncludePrivate,
			IncludeLabels:  opts.IncludeLabels,
		},
		rootDir:    rootDir,
		source:     parsed.Source,
		file:       rel,
		moduleName: moduleName,
		moduleKind: moduleKind,
		attrs:      attrs,
	}
	return FileResult{
		Path:       rel,
		ModuleName: moduleName,
		ModuleKind: moduleKind,
		Parse: ParseSummary{
			HasError:   parsed.HasError,
			HasMissing: parsed.HasMissing,
		},
		Symbols: ext.extract(parsed.Root),
	}, nil
}

func discoverFiles(opts Options) ([]fileCandidate, error) {
	if strings.TrimSpace(opts.Path) != "" {
		root := resolvePath(opts.RootDir, opts.Path)
		return collectFiles(root, opts)
	}
	cfg := opts.Config
	dirs := []struct {
		path string
		kind string
	}{
		{cfg.Src.Modules, "standard"},
		{cfg.Src.Classes, "class"},
		{cfg.Src.Forms, "form"},
		{cfg.Src.Workbook, "document"},
	}
	seen := make(map[string]bool)
	files := make([]fileCandidate, 0)
	for _, dir := range dirs {
		if strings.TrimSpace(dir.path) == "" {
			continue
		}
		collected, err := collectFiles(filepath.Join(opts.RootDir, dir.path), opts)
		if err != nil {
			return nil, err
		}
		for _, file := range collected {
			if dir.kind == "form" {
				file.moduleKind = formFileKind(opts.RootDir, cfg, file.path)
			} else {
				file.moduleKind = dir.kind
			}
			abs, err := filepath.Abs(file.path)
			if err != nil {
				return nil, err
			}
			if seen[abs] {
				continue
			}
			seen[abs] = true
			files = append(files, file)
		}
	}
	sort.Slice(files, func(i, j int) bool {
		return files[i].path < files[j].path
	})
	return files, nil
}

func collectFiles(root string, opts Options) ([]fileCandidate, error) {
	info, err := os.Stat(root)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, nil
		}
		return nil, err
	}
	if !info.IsDir() {
		ext := strings.ToLower(filepath.Ext(root))
		if ext != ".bas" && ext != ".cls" && ext != ".frm" {
			return nil, fmt.Errorf("unsupported source extension: %s", ext)
		}
		return []fileCandidate{{path: root, moduleKind: kindForPath(opts.RootDir, opts.Config, root)}}, nil
	}
	files := make([]fileCandidate, 0)
	err = filepath.WalkDir(root, func(path string, d os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		ext := strings.ToLower(filepath.Ext(path))
		if ext != ".bas" && ext != ".cls" && ext != ".frm" {
			return nil
		}
		if shouldSkipFormArtifact(opts.RootDir, opts.Config, path) {
			return nil
		}
		files = append(files, fileCandidate{path: path, moduleKind: kindForPath(opts.RootDir, opts.Config, path)})
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.Slice(files, func(i, j int) bool {
		return files[i].path < files[j].path
	})
	return files, nil
}

func shouldSkipFormArtifact(root string, cfg config.Config, path string) bool {
	if !strings.EqualFold(cfg.UserForm.CodeSource, "sidecar") || !strings.EqualFold(filepath.Ext(path), ".frm") {
		return false
	}
	formsRoot := filepath.Clean(filepath.Join(root, cfg.Src.Forms))
	cleanPath := filepath.Clean(path)
	if !isPathInsideRoot(cleanPath, formsRoot) {
		return false
	}
	sidecar := filepath.Join(formsRoot, "code", strings.TrimSuffix(filepath.Base(cleanPath), filepath.Ext(cleanPath))+".bas")
	_, err := os.Stat(sidecar)
	return err == nil
}

func kindForPath(root string, cfg config.Config, path string) string {
	ext := strings.ToLower(filepath.Ext(path))
	workbookRoot := filepath.Clean(filepath.Join(root, cfg.Src.Workbook))
	formsRoot := filepath.Clean(filepath.Join(root, cfg.Src.Forms))
	classesRoot := filepath.Clean(filepath.Join(root, cfg.Src.Classes))
	if isPathInsideRoot(path, workbookRoot) {
		return "document"
	}
	switch ext {
	case ".frm":
		return "form"
	case ".cls":
		return "class"
	}
	if isPathInsideRoot(path, filepath.Join(formsRoot, "code")) && matchingFormArtifact(formsRoot, path) {
		return "form"
	}
	if isPathInsideRoot(path, classesRoot) {
		return "class"
	}
	return "standard"
}

func formFileKind(root string, cfg config.Config, path string) string {
	if strings.EqualFold(filepath.Ext(path), ".bas") && matchingFormArtifact(filepath.Join(root, cfg.Src.Forms), path) {
		return "form"
	}
	return kindForPath(root, cfg, path)
}

func matchingFormArtifact(formsRoot, path string) bool {
	formName := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	_, err := os.Stat(filepath.Join(formsRoot, formName+".frm"))
	return err == nil
}

func (e *extractor) extract(root *tree_sitter.Node) []Symbol {
	if root == nil {
		return nil
	}
	moduleRange := vbaast.NodeRange(root)
	e.symbols = append(e.symbols, Symbol{
		Name:        e.moduleName,
		Kind:        "module",
		Module:      e.moduleName,
		File:        e.file,
		StartLine:   moduleRange.StartLine,
		StartColumn: moduleRange.StartColumn,
		EndLine:     moduleRange.EndLine,
		EndColumn:   moduleRange.EndColumn,
		StartByte:   moduleRange.StartByte,
		EndByte:     moduleRange.EndByte,
		Signature:   "Module " + e.moduleName,
		Attributes:  e.attrs,
	})
	for i := uint(0); i < root.NamedChildCount(); i++ {
		e.visit(root.NamedChild(i), "")
	}
	return e.symbols
}

func (e *extractor) visit(node *tree_sitter.Node, parentProc string) {
	if node == nil {
		return
	}
	switch node.Kind() {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration":
		sym := e.procedureSymbol(node)
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
		for _, param := range e.parameterSymbols(node, sym.Name) {
			if e.includeSymbol(param) {
				e.symbols = append(e.symbols, param)
			}
		}
		parentProc = sym.Name
	case "declare_statement", "declare_sub_statement", "declare_function_statement":
		sym := e.simpleSymbol(node, "declare", "")
		switch node.Kind() {
		case "declare_sub_statement":
			sym.Kind = "declare_sub"
		case "declare_function_statement":
			sym.Kind = "declare_function"
		default:
			if hasWord(sym.Signature, "Declare") && hasWord(sym.Signature, "Sub") {
				sym.Kind = "declare_sub"
			} else if hasWord(sym.Signature, "Declare") && hasWord(sym.Signature, "Function") {
				sym.Kind = "declare_function"
			}
		}
		sym.ReturnType = typeText(node, e.source)
		sym.Parameters = parameters(node, e.source)
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	case "type_declaration":
		sym := e.simpleSymbol(node, "type", "")
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	case "enum_declaration":
		sym := e.simpleSymbol(node, "enum", "")
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	case "const_declaration":
		e.constSymbols(node, parentProc)
	case "variable_declaration":
		e.variableSymbols(node, parentProc)
	case "implements_statement":
		sym := e.implementsSymbol(node)
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	case "label_statement":
		if e.opts.IncludeLabels {
			sym := e.simpleSymbol(node, "label", parentProc)
			e.symbols = append(e.symbols, sym)
		}
	case "line_number_statement":
		if e.opts.IncludeLabels {
			sym := e.lineNumberSymbol(node, parentProc)
			e.symbols = append(e.symbols, sym)
		}
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		if child.Kind() == "block" || node.Kind() == "source_file" || node.Kind() == "block" {
			e.visit(child, parentProc)
		}
	}
}

func (e *extractor) procedureSymbol(node *tree_sitter.Node) Symbol {
	kind := strings.TrimSuffix(node.Kind(), "_declaration")
	switch node.Kind() {
	case "property_get_declaration":
		kind = "property_get"
	case "property_let_declaration":
		kind = "property_let"
	case "property_set_declaration":
		kind = "property_set"
	case "property_declaration":
		sig := strings.ToLower(firstLine(node.Utf8Text(e.source)))
		switch {
		case strings.Contains(sig, "property get"):
			kind = "property_get"
		case strings.Contains(sig, "property let"):
			kind = "property_let"
		case strings.Contains(sig, "property set"):
			kind = "property_set"
		default:
			kind = "property"
		}
	}
	sym := e.simpleSymbol(node, kind, "")
	sym.Static = hasField(node, "static_modifier") || hasWord(sym.Signature, "Static")
	sym.ReturnType = typeText(node, e.source)
	sym.Parameters = parameters(node, e.source)
	return sym
}

func (e *extractor) constSymbols(node *tree_sitter.Node, parentProc string) {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() != "const_declarator" {
			continue
		}
		sym := e.symbolFromNode(child, "const", parentProc)
		sym.Visibility = visibilityText(node, e.source)
		sym.Signature = firstLine(node.Utf8Text(e.source))
		sym.ReturnType = typeText(child, e.source)
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	}
}

func (e *extractor) variableSymbols(node *tree_sitter.Node, parentProc string) {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil || child.Kind() != "variable_declarator" {
			continue
		}
		kind := "module_variable"
		if parentProc != "" {
			kind = "local_variable"
		} else if e.moduleKind == "class" || e.moduleKind == "form" {
			kind = "field"
		}
		sym := e.symbolFromNode(child, kind, parentProc)
		sym.Visibility = visibilityText(node, e.source)
		sym.Signature = firstLine(node.Utf8Text(e.source))
		sym.ReturnType = typeText(child, e.source)
		sym.Static = hasField(node, "static_modifier") || hasWord(sym.Signature, "Static")
		if hasField(node, "with_events_modifier") || hasWord(sym.Signature, "WithEvents") {
			sym.Kind = "withevents_field"
		}
		if e.includeSymbol(sym) {
			e.symbols = append(e.symbols, sym)
		}
	}
}

func (e *extractor) parameterSymbols(node *tree_sitter.Node, parentProc string) []Symbol {
	list := node.ChildByFieldName("parameters")
	if list == nil {
		list = firstNamedChildKind(node, "parameter_list")
	}
	if list == nil {
		return nil
	}
	out := make([]Symbol, 0, list.NamedChildCount())
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil || child.Kind() != "parameter" {
			continue
		}
		sym := e.parameterSymbol(child, parentProc)
		if strings.TrimSpace(sym.Name) != "" {
			out = append(out, sym)
		}
	}
	return out
}

func (e *extractor) parameterSymbol(node *tree_sitter.Node, parentProc string) Symbol {
	r := vbaast.NodeRange(node)
	nameNode := node.ChildByFieldName("name")
	if nameNode == nil {
		nameNode = firstNamedChildKind(node, "identifier")
	}
	if nameNode != nil {
		r = vbaast.NodeRange(nameNode)
	}
	name := nodeName(node, e.source)
	return Symbol{
		Name:        name,
		Kind:        "parameter",
		Visibility:  "",
		Module:      e.moduleName,
		File:        e.file,
		Parent:      parentProc,
		StartLine:   r.StartLine,
		StartColumn: r.StartColumn,
		EndLine:     r.EndLine,
		EndColumn:   r.EndColumn,
		StartByte:   r.StartByte,
		EndByte:     r.EndByte,
		Signature:   firstLine(node.Utf8Text(e.source)),
		ReturnType:  typeText(node, e.source),
	}
}

func (e *extractor) implementsSymbol(node *tree_sitter.Node) Symbol {
	name := ""
	if nameNode := node.ChildByFieldName("name"); nameNode != nil {
		name = nameNode.Utf8Text(e.source)
	}
	if name == "" {
		name = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(node.Utf8Text(e.source)), "Implements"))
	}
	sym := e.simpleSymbol(node, "implements", "")
	sym.Name = name
	if sym.Name == "" {
		sym.Name = node.Utf8Text(e.source)
	}
	return sym
}

func (e *extractor) lineNumberSymbol(node *tree_sitter.Node, parentProc string) Symbol {
	sym := e.simpleSymbol(node, "line_number_label", parentProc)
	if number := node.ChildByFieldName("number"); number != nil {
		sym.Name = number.Utf8Text(e.source)
	} else if number := firstNamedChildKind(node, "line_number_literal"); number != nil {
		sym.Name = number.Utf8Text(e.source)
	}
	return sym
}

func (e *extractor) simpleSymbol(node *tree_sitter.Node, kind, parent string) Symbol {
	return e.symbolFromNode(node, kind, parent)
}

func (e *extractor) symbolFromNode(node *tree_sitter.Node, kind, parent string) Symbol {
	r := vbaast.NodeRange(node)
	name := nodeName(node, e.source)
	sig := firstLine(node.Utf8Text(e.source))
	return Symbol{
		Name:        name,
		Kind:        kind,
		Visibility:  visibilityText(node, e.source),
		Module:      e.moduleName,
		File:        e.file,
		Parent:      parent,
		StartLine:   r.StartLine,
		StartColumn: r.StartColumn,
		EndLine:     r.EndLine,
		EndColumn:   r.EndColumn,
		StartByte:   r.StartByte,
		EndByte:     r.EndByte,
		Signature:   sig,
	}
}

func (e *extractor) includeSymbol(sym Symbol) bool {
	if sym.Kind == "module" || sym.Kind == "attribute" || sym.Kind == "implements" {
		return true
	}
	if sym.Kind == "local_variable" || sym.Kind == "parameter" || (sym.Kind == "const" && sym.Parent != "") {
		return e.opts.IncludePrivate
	}
	if e.opts.IncludePrivate {
		return true
	}
	if sym.Visibility == "" {
		switch sym.Kind {
		case "sub", "function", "property_get", "property_let", "property_set", "property", "declare", "declare_sub", "declare_function", "type", "enum", "const":
			return true
		default:
			return false
		}
	}
	return !strings.EqualFold(sym.Visibility, "Private") && sym.Visibility != ""
}

func moduleMetadata(path string, source []byte) (string, []Attribute) {
	attrs := make([]Attribute, 0)
	moduleName := ""
	for _, line := range strings.Split(string(source), "\n") {
		match := attrRe.FindStringSubmatch(line)
		if match == nil {
			continue
		}
		value := strings.Trim(strings.TrimSpace(match[2]), `"`)
		attr := Attribute{Name: match[1], Value: value}
		attrs = append(attrs, attr)
		if strings.EqualFold(attr.Name, "VB_Name") && attr.Value != "" {
			moduleName = attr.Value
		}
	}
	if moduleName == "" {
		moduleName = strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
	}
	return moduleName, attrs
}

func nodeName(node *tree_sitter.Node, source []byte) string {
	if name := node.ChildByFieldName("name"); name != nil {
		return name.Utf8Text(source)
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return name.Utf8Text(source)
	}
	if name := firstNamedChildKind(node, "line_number_literal"); name != nil {
		return name.Utf8Text(source)
	}
	return ""
}

func firstNamedChildKind(node *tree_sitter.Node, kind string) *tree_sitter.Node {
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child != nil && child.Kind() == kind {
			return child
		}
	}
	return nil
}

func visibilityText(node *tree_sitter.Node, source []byte) string {
	if visibility := node.ChildByFieldName("visibility"); visibility != nil {
		return normalizeKeyword(visibility.Utf8Text(source))
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		child := node.NamedChild(i)
		if child == nil {
			continue
		}
		if child.Kind() == "visibility" {
			return normalizeKeyword(child.Utf8Text(source))
		}
		if child.Kind() == "procedure_modifier" {
			for j := uint(0); j < child.NamedChildCount(); j++ {
				mod := child.NamedChild(j)
				if mod != nil && mod.Kind() == "visibility" {
					return normalizeKeyword(mod.Utf8Text(source))
				}
			}
		}
	}
	text := firstLine(node.Utf8Text(source))
	for _, word := range []string{"Public", "Private", "Friend"} {
		if hasWord(text, word) {
			return word
		}
	}
	return ""
}

func typeText(node *tree_sitter.Node, source []byte) string {
	asType := node.ChildByFieldName("type")
	if asType == nil {
		asType = firstNamedChildKind(node, "as_type_clause")
	}
	if asType == nil {
		return ""
	}
	if typeExpr := asType.ChildByFieldName("type"); typeExpr != nil {
		return strings.TrimSpace(typeExpr.Utf8Text(source))
	}
	if asType.Kind() == "type_expression" {
		return strings.TrimSpace(asType.Utf8Text(source))
	}
	text := strings.TrimSpace(asType.Utf8Text(source))
	if strings.HasPrefix(strings.ToLower(text), "as ") {
		return strings.TrimSpace(text[3:])
	}
	return text
}

func parameters(node *tree_sitter.Node, source []byte) []Parameter {
	list := node.ChildByFieldName("parameters")
	if list == nil {
		list = firstNamedChildKind(node, "parameter_list")
	}
	if list == nil {
		return nil
	}
	params := make([]Parameter, 0)
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil || child.Kind() != "parameter" {
			continue
		}
		param := Parameter{Name: nodeName(child, source), Type: typeText(child, source)}
		if passing := child.ChildByFieldName("passing_mode"); passing != nil {
			param.Passing = modifierKeyword(passing)
		} else {
			text := child.Utf8Text(source)
			if hasWord(text, "ByVal") {
				param.Passing = "ByVal"
			} else if hasWord(text, "ByRef") {
				param.Passing = "ByRef"
			}
		}
		param.Optional = hasField(child, "optional_modifier") || hasWord(child.Utf8Text(source), "Optional")
		initializer := child.ChildByFieldName("default_value")
		if initializer == nil {
			initializer = firstNamedChildKind(child, "initializer")
		}
		if initializer != nil {
			value := initializerText(initializer, source)
			param.Default = &value
		}
		params = append(params, param)
	}
	return params
}

func hasField(node *tree_sitter.Node, field string) bool {
	return node != nil && node.ChildByFieldName(field) != nil
}

func initializerText(node *tree_sitter.Node, source []byte) string {
	if value := node.ChildByFieldName("value"); value != nil {
		return strings.TrimSpace(value.Utf8Text(source))
	}
	return strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(node.Utf8Text(source)), "="))
}

func modifierKeyword(node *tree_sitter.Node) string {
	if node == nil {
		return ""
	}
	switch node.Kind() {
	case "byval_modifier":
		return "ByVal"
	case "byref_modifier":
		return "ByRef"
	}
	return normalizeKeyword(strings.TrimSuffix(node.Kind(), "_modifier"))
}

func firstLine(text string) string {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	for _, line := range strings.Split(text, "\n") {
		if strings.TrimSpace(line) != "" {
			return strings.TrimSpace(line)
		}
	}
	return ""
}

func hasWord(text, word string) bool {
	fields := strings.FieldsFunc(text, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	for _, field := range fields {
		if strings.EqualFold(field, word) {
			return true
		}
	}
	return false
}

func isVBAIdentifierRune(r rune) bool {
	switch r {
	case '_', '$', '%', '&', '!', '#', '@', '^':
		return true
	}
	return r >= '0' && r <= '9' || r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z'
}

func normalizeKeyword(text string) string {
	if text == "" {
		return ""
	}
	lower := strings.ToLower(strings.TrimSpace(text))
	return strings.ToUpper(lower[:1]) + lower[1:]
}

func displayPath(rootDir, path string) string {
	rel, err := filepath.Rel(rootDir, path)
	if err != nil {
		return filepath.ToSlash(path)
	}
	return filepath.ToSlash(rel)
}

func resolvePath(root, path string) string {
	if filepath.IsAbs(path) {
		return path
	}
	return filepath.Join(root, path)
}

func isPathInsideRoot(path, root string) bool {
	path = filepath.Clean(path)
	root = filepath.Clean(root)
	if strings.EqualFold(path, root) {
		return true
	}
	rel, err := filepath.Rel(root, path)
	if err != nil {
		return false
	}
	return rel != "." && !strings.HasPrefix(rel, "..") && !filepath.IsAbs(rel)
}
