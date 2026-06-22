package analyze

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/suppression"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Finding struct {
	Code       string   `json:"code"`
	Severity   string   `json:"severity"`
	File       string   `json:"file"`
	Module     string   `json:"module,omitempty"`
	Procedure  string   `json:"procedure,omitempty"`
	Line       int      `json:"line"`
	Column     int      `json:"column,omitempty"`
	Message    string   `json:"message"`
	Reason     string   `json:"reason"`
	Suggestion string   `json:"suggestion"`
	NearbyCode []string `json:"nearby_code,omitempty"`
}

type Result struct {
	Findings []Finding
	Warnings []map[string]any
}

type Analyzer struct {
	RootDir    string
	Config     config.Config
	PathFilter func(string) bool
}

var (
	declRe             = regexp.MustCompile(`(?i)^\s*(?:dim|private|public|static)\s+(.+)$`)
	assignRe           = regexp.MustCompile(`(?i)^\s*(?:let\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	setAssignRe        = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	callAssignRe       = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*(?:\(|$)`)
	withRe             = regexp.MustCompile(`(?i)^\s*with\s+(.+)$`)
	endWithRe          = regexp.MustCompile(`(?i)^\s*end\s+with\b`)
	withMemberRe       = regexp.MustCompile(`(?i)^\s*\.([A-Za-z_][A-Za-z0-9_]*)\b`)
	memberRe           = regexp.MustCompile(`(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\b`)
	traceHelperCallRe  = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(XlflowLog|XlflowSetTraceFile)\b`)
	traceHelperQualRe  = regexp.MustCompile(`(?i)\bXlflowTrace\s*\.\s*(XlflowLog|XlflowSetTraceFile)\b`)
	unqualifiedExcelRe = regexp.MustCompile(`(?i)(^|[^A-Za-z0-9_.$])\b(Range|Cells|Rows|Columns)\b\s*(?:\(|\.)`)
	activeExcelRe      = regexp.MustCompile(`(?i)(^|[^A-Za-z0-9_.$])\b(ActiveWorkbook|ActiveSheet|ActiveCell|Selection)\b`)
	redimPreserveRe    = regexp.MustCompile(`(?i)^\s*redim\s+preserve\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)`)
)

var objectTypes = map[string]bool{
	"application": true, "workbook": true, "worksheet": true, "range": true,
	"chart": true, "pivot table": true, "pivottable": true, "listobject": true,
	"dictionary": true, "collection": true, "object": true, "window": true,
}

type invalidMemberRule struct {
	Code       string
	Reason     string
	Suggestion string
}

type helperDependencyRule struct {
	Code       string
	Reason     string
	Suggestion string
}

var invalidObjectMembers = map[string]map[string]invalidMemberRule{
	"worksheet": {
		"displaygridlines": {
			Code:       "VBA104",
			Reason:     "DisplayGridlines is a Window property, not a Worksheet member.",
			Suggestion: "Set DisplayGridlines on ActiveWindow or another Window object instead of the Worksheet.",
		},
		"screenupdating": {
			Code:       "VBA211",
			Reason:     "ScreenUpdating is an Application property, not a Worksheet member.",
			Suggestion: "Set Application.ScreenUpdating instead of a Worksheet member.",
		},
		"displayalerts": {
			Code:       "VBA211",
			Reason:     "DisplayAlerts is an Application property, not a Worksheet member.",
			Suggestion: "Set Application.DisplayAlerts instead of a Worksheet member.",
		},
		"enableevents": {
			Code:       "VBA211",
			Reason:     "EnableEvents is an Application property, not a Worksheet member.",
			Suggestion: "Set Application.EnableEvents instead of a Worksheet member.",
		},
	},
	"workbook": {
		"displaygridlines": {
			Code:       "VBA211",
			Reason:     "DisplayGridlines is a Window property, not a Workbook member.",
			Suggestion: "Set DisplayGridlines on ActiveWindow or another Window object instead of the Workbook.",
		},
		"screenupdating": {
			Code:       "VBA211",
			Reason:     "ScreenUpdating is an Application property, not a Workbook member.",
			Suggestion: "Set Application.ScreenUpdating instead of a Workbook member.",
		},
	},
}

var traceHelperDependencies = map[string]helperDependencyRule{
	"xlflowlog": {
		Code:       "VBA105",
		Reason:     "XlflowLog belongs to the removed xlflow trace workflow and is no longer a supported debugging API.",
		Suggestion: "Replace `XlflowLog` with `XlflowDebug.Log`, then inspect debug output through `xlflow run --json` or `xlflow test --json`.",
	},
	"xlflowsettracefile": {
		Code:       "VBA106",
		Reason:     "XlflowSetTraceFile belongs to the removed xlflow trace workflow and should not be called from project VBA.",
		Suggestion: "Delete `XlflowSetTraceFile` calls and emit runtime diagnostics with `XlflowDebug.Log` instead. `xlflow run --json` is the supported machine-readable execution surface.",
	},
}

type analysisContext struct {
	functionReturns map[string]string
	procedures      map[string]procedureSignature
}

type procedureSignature struct {
	Name       string
	ReturnType string
	Params     []parameterInfo
}

type parameterInfo struct {
	Name    string
	Type    string
	Passing string
}

type parsedFile struct {
	Path   string
	Lines  []string
	Module string
	Root   *tree_sitter.Node
	Result *vbaast.ParseResult
}

type sourceProcedure struct {
	Kind       string
	Name       string
	ReturnType string
	StartLine  int
	EndLine    int
	Params     []parameterInfo
}

type sourceDeclaration struct {
	Name          string
	Type          string
	Line          int
	Object        bool
	Array         bool
	NewExpression bool
	Parameter     bool
}

type withInfo struct {
	Target string
	Type   string
}

func (a Analyzer) Run() ([]Finding, error) {
	result, err := a.RunResult()
	if err != nil {
		return nil, err
	}
	return result.Findings, nil
}

func (a Analyzer) RunResult() (Result, error) {
	files, err := a.files()
	if err != nil {
		return Result{}, err
	}
	parser, err := vbaast.NewParser()
	if err != nil {
		return Result{}, err
	}
	defer parser.Close()
	parsedFiles := make([]parsedFile, 0, len(files))
	for _, file := range files {
		parsed, err := parser.ParseFile(file)
		if err != nil {
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		if parsed.HasError || parsed.HasMissing {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, fmt.Errorf("parse %s: VBA parser reported errors or missing nodes", file)
		}
		parsedFiles = append(parsedFiles, parsedFile{
			Path:   file,
			Lines:  normalizedSourceLines(string(parsed.Source)),
			Module: strings.TrimSuffix(filepath.Base(file), filepath.Ext(file)),
			Root:   parsed.Root,
			Result: parsed,
		})
	}
	defer closeParsedFiles(parsedFiles)

	ctx := a.buildContext(parsedFiles)
	var findings []Finding
	for _, file := range parsedFiles {
		findings = append(findings, a.analyzeParsedFile(file, ctx)...)
	}
	sortFindings(findings)
	directives, warnings, err := suppression.DirectivesForFiles(a.RootDir, files)
	if err != nil {
		return Result{}, err
	}
	findings, suppressionWarnings := applyInlineSuppressions(findings, directives)
	warnings = append(warnings, suppressionWarnings...)
	return Result{Findings: findings, Warnings: warnings}, nil
}

func closeParsedFiles(files []parsedFile) {
	for _, file := range files {
		if file.Result != nil {
			file.Result.Close()
		}
	}
}

func (a Analyzer) files() ([]string, error) {
	dirs := []string{a.Config.Src.Modules, a.Config.Src.Classes, a.Config.Src.Forms, a.Config.Src.Workbook, "tests"}
	var files []string
	for _, dir := range dirs {
		root := filepath.Join(a.RootDir, dir)
		if _, err := os.Stat(root); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return nil, err
		}
		err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if d.IsDir() {
				return nil
			}
			switch strings.ToLower(filepath.Ext(path)) {
			case ".bas", ".cls", ".frm":
				if !a.shouldIncludeFile(path) {
					return nil
				}
				files = append(files, path)
			}
			return nil
		})
		if err != nil {
			return nil, err
		}
	}
	sort.Strings(files)
	return files, nil
}

func (a Analyzer) shouldIncludeFile(path string) bool {
	if a.PathFilter != nil && !a.PathFilter(path) {
		return false
	}
	if !strings.EqualFold(filepath.Ext(path), ".frm") {
		return true
	}
	if !strings.EqualFold(a.Config.UserForm.CodeSource, "sidecar") {
		return true
	}
	formsRoot := filepath.Clean(filepath.Join(a.RootDir, a.Config.Src.Forms))
	cleanPath := filepath.Clean(path)
	if !isPathInsideRoot(cleanPath, formsRoot) {
		return true
	}
	sidecarPath := filepath.Join(formsRoot, "code", strings.TrimSuffix(filepath.Base(cleanPath), filepath.Ext(cleanPath))+".bas")
	if _, err := os.Stat(sidecarPath); err == nil {
		return false
	}
	return true
}

func (a Analyzer) buildContext(files []parsedFile) analysisContext {
	ctx := analysisContext{functionReturns: map[string]string{}, procedures: map[string]procedureSignature{}}
	for _, file := range files {
		for _, proc := range sourceProceduresFromAST(file.Root, file.Result.Source) {
			if isObjectType(proc.ReturnType) {
				ctx.functionReturns[strings.ToLower(proc.Name)] = proc.ReturnType
			}
			ctx.procedures[strings.ToLower(proc.Name)] = procedureSignature{
				Name:       proc.Name,
				ReturnType: proc.ReturnType,
				Params:     proc.Params,
			}
			ctx.procedures[strings.ToLower(file.Module+"."+proc.Name)] = ctx.procedures[strings.ToLower(proc.Name)]
		}
	}
	return ctx
}

func (a Analyzer) analyzeParsedFile(file parsedFile, ctx analysisContext) []Finding {
	reportedMissingHelpers := map[string]bool{}
	var findings []Finding
	procedures := sourceProceduresFromAST(file.Root, file.Result.Source)
	moduleDecls := moduleDeclarations(file.Lines, procedures)
	for _, proc := range procedures {
		findings = append(findings, a.analyzeProcedure(file, proc, moduleDecls, ctx, reportedMissingHelpers)...)
	}
	if len(procedures) == 0 {
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		findings = append(findings, a.analyzeProcedure(file, proc, moduleDecls, ctx, reportedMissingHelpers)...)
	}
	return findings
}

func (a Analyzer) analyzeProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, reportedMissingHelpers map[string]bool) []Finding {
	decls := cloneDeclarations(moduleDecls)
	for key, decl := range procedureDeclarations(file.Lines, proc) {
		decls[key] = decl
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Line: proc.StartLine, Object: isObjectType(param.Type), Parameter: true}
	}
	handlerLabels := onErrorHandlerLabels(file.Lines, proc)
	withStack := make([]withInfo, 0)
	initialized := initialObjectState(decls)
	maybeInitializedByCall := map[string]bool{}
	findAssignments := map[string]int{}
	guardedFinds := map[string]bool{}
	functionAssigned := false
	var findings []Finding

	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		if stmt == "" {
			continue
		}
		lower := strings.ToLower(stmt)

		if endWithRe.MatchString(stmt) {
			if len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
			continue
		}
		if m := withRe.FindStringSubmatch(stmt); len(m) > 0 {
			withStack = append(withStack, resolveWithInfo(m[1], decls))
			continue
		}
		for _, helper := range referencedTraceHelpers(stmt) {
			key := strings.ToLower(helper)
			if reportedMissingHelpers[key] {
				continue
			}
			if rule, ok := traceHelperDependencies[key]; ok {
				findings = append(findings, a.helperFinding(file, proc, lineNo, helper, rule))
				reportedMissingHelpers[key] = true
			}
		}
		if a.Config.Analyze.DetectExcelObjectMemberMismatch {
			findings = append(findings, a.memberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)...)
		} else {
			findings = append(findings, a.legacyMemberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)...)
		}
		if setAssignRe.MatchString(stmt) {
			if a.Config.Analyze.DetectObjectUseBeforeSet {
				findings = append(findings, a.objectUseBeforeSetFindings(file, proc, lineNo, stmt, decls, initialized, maybeInitializedByCall)...)
			}
			if m := setAssignRe.FindStringSubmatch(stmt); len(m) > 0 {
				initialized[strings.ToLower(m[1])] = true
				if strings.EqualFold(m[1], proc.Name) {
					functionAssigned = true
				}
			}
			if name, ok := rangeFindAssignment(stmt); ok {
				findAssignments[strings.ToLower(name)] = lineNo
			}
			if a.Config.Analyze.ForbidUnqualifiedExcelObjects {
				findings = append(findings, a.unqualifiedExcelFindings(file, proc, lineNo, stmt)...)
			}
			continue
		}
		if m := assignRe.FindStringSubmatch(stmt); len(m) > 0 {
			target := strings.ToLower(m[1])
			if proc.Name != "" && strings.EqualFold(target, proc.Name) {
				functionAssigned = true
			}
			if proc.Name != "" && strings.EqualFold(target, proc.Name) && isObjectType(proc.ReturnType) {
				findings = append(findings, a.objectSetFinding(file, proc, lineNo, "VBA103", m[1], proc.ReturnType))
				continue
			}
			if cm := callAssignRe.FindStringSubmatch(stmt); len(cm) > 0 {
				callee := strings.ToLower(lastName(cm[2]))
				if typ, ok := decls[target]; ok && typ.Object && isObjectType(ctx.functionReturns[callee]) {
					findings = append(findings, a.objectSetFinding(file, proc, lineNo, "VBA102", m[1], ctx.functionReturns[callee]))
					continue
				}
			}
			if decl, ok := decls[target]; ok && decl.Object {
				findings = append(findings, a.objectSetFinding(file, proc, lineNo, "VBA101", m[1], decl.Type))
			}
		}
		if a.Config.Analyze.DetectRangeFindNothingCheck {
			findings = append(findings, a.rangeFindFindings(file, proc, lineNo, stmt, findAssignments, guardedFinds)...)
		}
		if a.Config.Analyze.DetectObjectUseBeforeSet {
			findings = append(findings, a.objectUseBeforeSetFindings(file, proc, lineNo, stmt, decls, initialized, maybeInitializedByCall)...)
		}
		if a.Config.Analyze.ForbidUnqualifiedExcelObjects {
			findings = append(findings, a.unqualifiedExcelFindings(file, proc, lineNo, stmt)...)
		}
		if a.Config.Analyze.DetectByRefArgumentMismatch {
			findings = append(findings, a.byRefMismatchFindings(file, proc, lineNo, stmt, ctx)...)
		}
		if a.Config.Analyze.DetectDictionaryCollectionGuard {
			findings = append(findings, a.dictionaryCollectionFindings(file, proc, lineNo, stmt, decls)...)
		}
		if a.Config.Analyze.DetectRedimPreserveDimension {
			findings = append(findings, a.redimPreserveFindings(file, proc, lineNo, stmt)...)
		}
		if a.Config.Analyze.DetectObjectArrayComparison {
			findings = append(findings, a.objectArrayComparisonFindings(file, proc, lineNo, stmt, decls)...)
		}
		markCallInitialized(stmt, decls, ctx, maybeInitializedByCall)
		_ = lower
	}
	if a.Config.Analyze.DetectApplicationStateRestore {
		findings = append(findings, a.applicationStateFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectErrorHandlerFallthrough {
		findings = append(findings, a.errorHandlerFallthroughFindings(file, proc, handlerLabels)...)
	}
	if a.Config.Analyze.DetectFunctionReturnPath && proc.Kind == "Function" && proc.Name != "" && !functionAssigned {
		findings = append(findings, a.simpleFinding(file, proc, proc.StartLine, "VBA210", "warning", proc.Name+" may exit without assigning its return value.", "Functions return the default value when no assignment to the function name is reached.", "Assign "+proc.Name+" on every successful return path, or make the default return explicit."))
	}
	return findings
}

func sourceProceduresFromAST(root *tree_sitter.Node, source []byte) []sourceProcedure {
	var procedures []sourceProcedure
	collectSourceProcedures(root, source, &procedures)
	sort.SliceStable(procedures, func(i, j int) bool {
		return procedures[i].StartLine < procedures[j].StartLine
	})
	return procedures
}

func collectSourceProcedures(node *tree_sitter.Node, source []byte, procedures *[]sourceProcedure) {
	if node == nil {
		return
	}
	switch node.Kind() {
	case "sub_declaration", "function_declaration", "property_declaration", "property_get_declaration", "property_let_declaration", "property_set_declaration":
		r := vbaast.NodeRange(node)
		kind := "Sub"
		if strings.Contains(node.Kind(), "function") {
			kind = "Function"
		} else if strings.Contains(node.Kind(), "property") {
			kind = "Property"
		}
		*procedures = append(*procedures, sourceProcedure{
			Kind:       kind,
			Name:       nodeProcedureName(node, source),
			ReturnType: typeText(node, source),
			StartLine:  r.StartLine,
			EndLine:    r.EndLine,
			Params:     parameters(node, source),
		})
		return
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		collectSourceProcedures(node.NamedChild(i), source, procedures)
	}
}

func nodeProcedureName(node *tree_sitter.Node, source []byte) string {
	if name := node.ChildByFieldName("name"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	return ""
}

func parameters(node *tree_sitter.Node, source []byte) []parameterInfo {
	list := node.ChildByFieldName("parameters")
	if list == nil {
		list = firstNamedChildKind(node, "parameter_list")
	}
	if list == nil {
		return nil
	}
	var params []parameterInfo
	for i := uint(0); i < list.NamedChildCount(); i++ {
		child := list.NamedChild(i)
		if child == nil || child.Kind() != "parameter" {
			continue
		}
		param := parameterInfo{Name: nodeName(child, source), Type: typeText(child, source)}
		if hasWord(child.Utf8Text(source), "ByVal") {
			param.Passing = "ByVal"
		} else {
			param.Passing = "ByRef"
		}
		params = append(params, param)
	}
	return params
}

func procedureDeclarations(lines []string, proc sourceProcedure) map[string]sourceDeclaration {
	decls := map[string]sourceDeclaration{}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if lineNo == proc.StartLine && isProcedureHeaderLine(lower) {
			continue
		}
		if !strings.HasPrefix(lower, "dim ") && !strings.HasPrefix(lower, "static ") && !strings.HasPrefix(lower, "private ") && !strings.HasPrefix(lower, "public ") {
			continue
		}
		m := declRe.FindStringSubmatch(stmt)
		if len(m) == 0 {
			continue
		}
		for _, part := range strings.Split(m[1], ",") {
			name, typ, array, newExpr := declarationNameAndType(part)
			if name == "" {
				continue
			}
			decls[strings.ToLower(name)] = sourceDeclaration{Name: name, Type: typ, Line: lineNo, Object: isObjectType(typ), Array: array, NewExpression: newExpr}
		}
	}
	return decls
}

func isProcedureHeaderLine(lower string) bool {
	lower = strings.TrimSpace(lower)
	for _, prefix := range []string{"sub ", "function ", "property ", "private sub ", "private function ", "private property ", "public sub ", "public function ", "public property ", "friend sub ", "friend function ", "friend property "} {
		if strings.HasPrefix(lower, prefix) {
			return true
		}
	}
	return false
}

func moduleDeclarations(lines []string, procedures []sourceProcedure) map[string]sourceDeclaration {
	decls := map[string]sourceDeclaration{}
	for i := 0; i < len(lines); i++ {
		lineNo := i + 1
		if lineInAnyProcedure(lineNo, procedures) {
			continue
		}
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if !strings.HasPrefix(lower, "dim ") && !strings.HasPrefix(lower, "static ") && !strings.HasPrefix(lower, "private ") && !strings.HasPrefix(lower, "public ") {
			continue
		}
		m := declRe.FindStringSubmatch(stmt)
		if len(m) == 0 {
			continue
		}
		for _, part := range strings.Split(m[1], ",") {
			name, typ, array, newExpr := declarationNameAndType(part)
			if name == "" {
				continue
			}
			decls[strings.ToLower(name)] = sourceDeclaration{Name: name, Type: typ, Line: lineNo, Object: isObjectType(typ), Array: array, NewExpression: newExpr}
		}
	}
	return decls
}

func lineInAnyProcedure(line int, procedures []sourceProcedure) bool {
	for _, proc := range procedures {
		if proc.StartLine <= line && line <= proc.EndLine {
			return true
		}
	}
	return false
}

func cloneDeclarations(decls map[string]sourceDeclaration) map[string]sourceDeclaration {
	clone := make(map[string]sourceDeclaration, len(decls))
	for key, decl := range decls {
		clone[key] = decl
	}
	return clone
}

func declarationNameAndType(text string) (string, string, bool, bool) {
	text = strings.TrimSpace(text)
	if text == "" {
		return "", "", false, false
	}
	lower := strings.ToLower(text)
	namePart := text
	typ := ""
	if idx := strings.Index(lower, " as "); idx >= 0 {
		namePart = strings.TrimSpace(text[:idx])
		typ = strings.TrimSpace(text[idx+4:])
	}
	newExpr := false
	if strings.HasPrefix(strings.ToLower(typ), "new ") {
		newExpr = true
		typ = strings.TrimSpace(typ[4:])
	}
	array := strings.Contains(namePart, "(") || strings.Contains(strings.ToLower(typ), "()")
	namePart = strings.TrimSpace(strings.TrimLeft(namePart, "()"))
	nameFields := strings.FieldsFunc(namePart, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	if len(nameFields) == 0 {
		return "", typ, array, newExpr
	}
	return cleanIdentifier(nameFields[0]), typ, array, newExpr
}

func initialObjectState(decls map[string]sourceDeclaration) map[string]bool {
	out := map[string]bool{}
	for key, decl := range decls {
		if decl.Object {
			out[key] = decl.Parameter || decl.NewExpression
		}
	}
	return out
}

func (a Analyzer) legacyMemberMismatchFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls map[string]sourceDeclaration, withStack []withInfo) []Finding {
	all := a.memberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)
	filtered := all[:0]
	for _, finding := range all {
		if finding.Code == "VBA104" {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func (a Analyzer) memberMismatchFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls map[string]sourceDeclaration, withStack []withInfo) []Finding {
	var findings []Finding
	if currentWith, ok := currentWithInfo(withStack); ok {
		if m := withMemberRe.FindStringSubmatch(stmt); len(m) > 0 {
			if rule, ok := invalidMemberRuleFor(currentWith.Type, m[1]); ok {
				findings = append(findings, a.memberFinding(file, proc, lineNo, currentWith.Target, currentWith.Type, m[1], rule))
			}
		}
	}
	for _, m := range memberRe.FindAllStringSubmatch(stmt, -1) {
		if decl, ok := decls[strings.ToLower(m[1])]; ok {
			if rule, ok := invalidMemberRuleFor(decl.Type, m[2]); ok {
				findings = append(findings, a.memberFinding(file, proc, lineNo, m[1], decl.Type, m[2], rule))
			}
		}
	}
	return findings
}

func (a Analyzer) rangeFindFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, findAssignments map[string]int, guarded map[string]bool) []Finding {
	lower := strings.ToLower(stmt)
	for name := range findAssignments {
		if strings.Contains(lower, "if "+name+" is nothing") || strings.Contains(lower, "if not "+name+" is nothing") {
			guarded[name] = true
		}
	}
	if name, ok := rangeFindAssignment(stmt); ok {
		findAssignments[strings.ToLower(name)] = lineNo
		return nil
	}
	var findings []Finding
	for name, assignLine := range findAssignments {
		if guarded[name] {
			continue
		}
		if strings.Contains(lower, name+".") {
			suggestion := "Add If " + name + " Is Nothing Then handling after the Find assignment."
			if assignLine == 0 {
				suggestion = "Check the Find result for Nothing before dereferencing it."
			}
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA201", "warning", "Range.Find result "+name+" is dereferenced before a Nothing check.", "Range.Find returns Nothing when no match is found, so dereferencing the result can raise runtime error 91.", suggestion))
			guarded[name] = true
		}
	}
	return findings
}

func (a Analyzer) objectUseBeforeSetFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls map[string]sourceDeclaration, initialized, maybeInitializedByCall map[string]bool) []Finding {
	var findings []Finding
	lower := strings.ToLower(stmt)
	for key, decl := range decls {
		if !decl.Object || initialized[key] || maybeInitializedByCall[key] || lineNo <= decl.Line {
			continue
		}
		if strings.Contains(lower, key+".") {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA202", "warning", decl.Name+" may be used before it is assigned with Set.", "Object variables are Nothing until assigned with Set; member access before initialization can raise runtime error 91.", "Assign `Set "+decl.Name+" = ...` before using members, or guard `If "+decl.Name+" Is Nothing Then`."))
			initialized[key] = true
		}
	}
	return findings
}

func markCallInitialized(stmt string, decls map[string]sourceDeclaration, ctx analysisContext, maybeInitialized map[string]bool) {
	if strings.Contains(stmt, "=") {
		return
	}
	name, args, ok := parseSimpleCall(stmt)
	if !ok {
		return
	}
	sig, ok := ctx.procedures[strings.ToLower(name)]
	if !ok {
		return
	}
	for i, arg := range args {
		if i >= len(sig.Params) {
			break
		}
		param := sig.Params[i]
		if strings.EqualFold(param.Passing, "ByVal") || !isObjectType(param.Type) {
			continue
		}
		key := strings.ToLower(cleanIdentifier(arg))
		if decl, ok := decls[key]; ok && decl.Object {
			maybeInitialized[key] = true
		}
	}
}

func (a Analyzer) unqualifiedExcelFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string) []Finding {
	var findings []Finding
	for _, m := range unqualifiedExcelRe.FindAllStringSubmatch(stmt, -1) {
		name := m[2]
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Unqualified "+name+" access depends on the active worksheet.", "Unqualified Excel object access is resolved through the active sheet or selection at runtime.", "Qualify "+name+" with an explicit Worksheet or Range object."))
	}
	for _, m := range activeExcelRe.FindAllStringSubmatch(stmt, -1) {
		name := m[2]
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", name+" creates an active Excel object dependency.", "ActiveWorkbook, ActiveSheet, ActiveCell, and Selection depend on the user's current Excel state.", "Pass or capture explicit Workbook, Worksheet, or Range objects instead."))
	}
	return findings
}

func (a Analyzer) byRefMismatchFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, ctx analysisContext) []Finding {
	name, args, ok := parseSimpleCall(stmt)
	if !ok {
		return nil
	}
	sig, ok := ctx.procedures[strings.ToLower(name)]
	if !ok {
		return nil
	}
	var findings []Finding
	for i, arg := range args {
		if i >= len(sig.Params) {
			break
		}
		param := sig.Params[i]
		if strings.EqualFold(param.Passing, "ByVal") || param.Type == "" {
			continue
		}
		if obviousArgumentMismatch(arg, param.Type) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA206", "warning", "Argument for ByRef parameter "+param.Name+" may not match "+param.Type+".", "VBA ByRef arguments must be type-compatible with the declared parameter.", "Pass a variable of type "+param.Type+" or change the procedure parameter to ByVal when mutation is not required."))
		}
	}
	return findings
}

func (a Analyzer) dictionaryCollectionFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls map[string]sourceDeclaration) []Finding {
	lower := strings.ToLower(stmt)
	if strings.Contains(lower, ".exists(") || strings.Contains(lower, "on error") {
		return nil
	}
	var findings []Finding
	for key, decl := range decls {
		typ := strings.ToLower(cleanIdentifier(decl.Type))
		if typ != "dictionary" && typ != "collection" && typ != "scripting.dictionary" {
			continue
		}
		if strings.Contains(lower, key+"(") || strings.Contains(lower, key+".item(") {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA207", "warning", decl.Name+" item access has no obvious existence guard.", "Dictionary and Collection item lookup can fail at runtime when the key is missing.", "Guard the access with Exists, explicit error handling, or a prior key validation path."))
		}
	}
	return findings
}

func (a Analyzer) redimPreserveFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string) []Finding {
	m := redimPreserveRe.FindStringSubmatch(stmt)
	if len(m) == 0 || !strings.Contains(m[2], ",") {
		return nil
	}
	return []Finding{a.simpleFinding(file, proc, lineNo, "VBA208", "warning", "ReDim Preserve is used on a multi-dimensional array.", "VBA can only Preserve the last dimension of an array; changing earlier dimensions raises a runtime error.", "Only change the last dimension, or copy values into a newly sized array explicitly.")}
}

func (a Analyzer) objectArrayComparisonFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls map[string]sourceDeclaration) []Finding {
	lower := strings.ToLower(stmt)
	var findings []Finding
	for key, decl := range decls {
		if decl.Object && strings.Contains(lower, key+" = nothing") {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA209", "warning", decl.Name+" is compared to Nothing with =.", "Object references must be compared with Is Nothing, not the scalar equality operator.", "Use `If "+decl.Name+" Is Nothing Then` or `If Not "+decl.Name+" Is Nothing Then`."))
		}
		if decl.Array && identifierComparedAsOperand(lower, key) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA209", "warning", decl.Name+" appears to be compared as a scalar value.", "VBA arrays cannot be compared directly to scalar values.", "Compare explicit elements or bounds instead of the array variable itself."))
		}
	}
	return findings
}

func identifierComparedAsOperand(stmt, name string) bool {
	name = strings.ToLower(cleanIdentifier(name))
	if name == "" {
		return false
	}
	for i := 0; i < len(stmt); i++ {
		opLen := comparisonOperatorLength(stmt, i)
		if opLen == 0 {
			continue
		}
		left := stmt[:i]
		right := stmt[i+opLen:]
		if operandEndsWithBareIdentifier(left, name) || operandStartsWithBareIdentifier(right, name) {
			return true
		}
		i += opLen - 1
	}
	return false
}

func comparisonOperatorLength(stmt string, index int) int {
	if index < 0 || index >= len(stmt) {
		return 0
	}
	if index+1 < len(stmt) {
		switch stmt[index : index+2] {
		case "<>", "<=", ">=":
			return 2
		}
	}
	switch stmt[index] {
	case '=', '<', '>':
		return 1
	default:
		return 0
	}
}

func operandEndsWithBareIdentifier(text, name string) bool {
	fields := identifierFields(text)
	if len(fields) == 0 {
		return false
	}
	return fieldMatchesBareIdentifier(text, fields[len(fields)-1], name)
}

func operandStartsWithBareIdentifier(text, name string) bool {
	fields := identifierFields(text)
	if len(fields) == 0 {
		return false
	}
	return fieldMatchesBareIdentifier(text, fields[0], name)
}

type identifierField struct {
	Text       string
	Start, End int
}

func identifierFields(text string) []identifierField {
	if strings.TrimSpace(text) == "" {
		return nil
	}
	var fields []identifierField
	start := -1
	for i, r := range text {
		if isVBAIdentifierRune(r) {
			if start < 0 {
				start = i
			}
			continue
		}
		if start >= 0 {
			fields = append(fields, identifierField{Text: text[start:i], Start: start, End: i})
			start = -1
		}
	}
	if start >= 0 {
		fields = append(fields, identifierField{Text: text[start:], Start: start, End: len(text)})
	}
	return fields
}

func fieldMatchesBareIdentifier(text string, field identifierField, name string) bool {
	if strings.ToLower(cleanIdentifier(field.Text)) != name {
		return false
	}
	if previousNonSpace(text, field.Start) == '.' {
		return false
	}
	next := nextNonSpace(text, field.End)
	return next != '(' && next != '.'
}

func previousNonSpace(text string, index int) byte {
	for i := index - 1; i >= 0; i-- {
		if text[i] != ' ' && text[i] != '\t' {
			return text[i]
		}
	}
	return 0
}

func nextNonSpace(text string, index int) byte {
	for i := index; i < len(text); i++ {
		if text[i] != ' ' && text[i] != '\t' {
			return text[i]
		}
	}
	return 0
}

func (a Analyzer) applicationStateFindings(file parsedFile, proc sourceProcedure) []Finding {
	props := []string{"enableevents", "displayalerts", "screenupdating", "calculation"}
	var findings []Finding
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		stmt := normalizedCodeLine(file.Lines[i])
		lower := compactStatement(strings.ToLower(stmt))
		for _, prop := range props {
			prefix := "application." + prop + "="
			if !strings.Contains(lower, prefix) {
				continue
			}
			if prop != "calculation" && !strings.Contains(lower, prefix+"false") {
				continue
			}
			if prop == "calculation" && strings.Contains(lower, prefix+"xlcalculationautomatic") {
				continue
			}
			if hasLaterApplicationRestore(file.Lines, proc, i+1, prop) {
				continue
			}
			name := applicationStateName(prop)
			findings = append(findings, a.simpleFinding(file, proc, i+1, "VBA203", "warning", "Application."+name+" is changed without an obvious restore path.", "Macros that leave Excel application state changed can break later user or automation workflows after failure.", "Save the previous Application."+name+" value and restore it in a cleanup path."))
		}
	}
	return findings
}

func hasLaterApplicationRestore(lines []string, proc sourceProcedure, start int, prop string) bool {
	prefix := "application." + prop + "="
	for i := start; i < proc.EndLine && i < len(lines); i++ {
		lower := compactStatement(strings.ToLower(normalizedCodeLine(lines[i])))
		if !strings.Contains(lower, prefix) {
			continue
		}
		if prop == "calculation" || !strings.Contains(lower, prefix+"false") {
			return true
		}
	}
	return false
}

func (a Analyzer) errorHandlerFallthroughFindings(file parsedFile, proc sourceProcedure, handlerLabels map[string]bool) []Finding {
	if len(handlerLabels) == 0 {
		return nil
	}
	var findings []Finding
	lastCode := ""
	for i := proc.StartLine - 1; i < proc.EndLine-1 && i < len(file.Lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		if stmt == "" {
			continue
		}
		if label, ok := labelName(stmt); ok && handlerLabels[strings.ToLower(label)] {
			if !isCleanupFallthroughLabel(label) && !terminatesNormalFlow(lastCode) {
				findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA204", "warning", "Normal execution can fall through into error handler "+label+".", "Without Exit Sub, Exit Function, or Exit Property before the handler label, successful execution can run error handling code.", errorHandlerFallthroughSuggestion(proc, label)))
			}
		}
		lastCode = stmt
	}
	return findings
}

func onErrorHandlerLabels(lines []string, proc sourceProcedure) map[string]bool {
	labels := map[string]bool{}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(lines); i++ {
		stmt := normalizedCodeLine(lines[i])
		lower := strings.ToLower(stmt)
		if strings.HasPrefix(lower, "on error goto ") && lower != "on error goto 0" {
			label := cleanIdentifier(strings.TrimSpace(stmt[len("on error goto "):]))
			if label != "" {
				labels[strings.ToLower(label)] = true
			}
		}
	}
	return labels
}

func BlockingFindings(findings []Finding) []Finding {
	blocking := make([]Finding, 0)
	for _, finding := range findings {
		if strings.EqualFold(finding.Severity, "error") {
			blocking = append(blocking, finding)
		}
	}
	return blocking
}

func (a Analyzer) objectSetFinding(file parsedFile, proc sourceProcedure, line int, code, target, typ string) Finding {
	msg := target + " is declared As " + typ + " and is assigned without Set."
	reason := "VBA object references require `Set` when assigning an object value."
	suggestion := "Use `Set " + target + " = ...` when the right-hand side returns an object."
	if code == "VBA103" {
		msg = target + " returns As " + typ + " and is assigned without Set."
		reason = "Object-returning VBA functions must assign the function name with `Set` before returning."
		suggestion = "Use `Set " + target + " = ...` inside this function body when returning a " + typ + "."
	}
	return a.simpleFinding(file, proc, line, code, "warning", msg, reason, suggestion)
}

func (a Analyzer) memberFinding(file parsedFile, proc sourceProcedure, line int, target, typ, member string, rule invalidMemberRule) Finding {
	targetLabel := target
	if targetLabel == "" {
		targetLabel = "This object"
	}
	return a.simpleFinding(file, proc, line, rule.Code, "error", targetLabel+" is declared As "+typ+" but ."+member+" is not a member of "+typ+".", rule.Reason, rule.Suggestion)
}

func (a Analyzer) helperFinding(file parsedFile, proc sourceProcedure, line int, helper string, rule helperDependencyRule) Finding {
	return a.simpleFinding(file, proc, line, rule.Code, "error", helper+" is a removed legacy trace API and should not appear in project VBA.", rule.Reason, rule.Suggestion)
}

func (a Analyzer) simpleFinding(file parsedFile, proc sourceProcedure, line int, code, severity, message, reason, suggestion string) Finding {
	rel, err := filepath.Rel(a.RootDir, file.Path)
	if err != nil {
		rel = file.Path
	}
	return Finding{
		Code:       code,
		Severity:   severity,
		File:       filepath.ToSlash(rel),
		Module:     file.Module,
		Procedure:  proc.Name,
		Line:       line,
		Message:    message,
		Reason:     reason,
		Suggestion: suggestion,
		NearbyCode: nearby(file.Lines, line, 2),
	}
}

func nearby(lines []string, line, radius int) []string {
	start := line - radius
	if start < 1 {
		start = 1
	}
	end := line + radius
	if end > len(lines) {
		end = len(lines)
	}
	out := make([]string, 0, end-start+1)
	for i := start; i <= end; i++ {
		prefix := "  "
		if i == line {
			prefix = "> "
		}
		out = append(out, prefix+strconvItoa(i)+" | "+lines[i-1])
	}
	return out
}

func normalizedSourceLines(source string) []string {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	return strings.Split(source, "\n")
}

func normalizedCodeLine(line string) string {
	return strings.Join(strings.Fields(maskStringLiterals(gui.StripComment(line))), " ")
}

func maskStringLiterals(line string) string {
	var b strings.Builder
	b.Grow(len(line))
	inString := false
	for i := 0; i < len(line); i++ {
		if line[i] != '"' {
			if inString {
				b.WriteByte(' ')
			} else {
				b.WriteByte(line[i])
			}
			continue
		}
		b.WriteByte('"')
		if inString && i+1 < len(line) && line[i+1] == '"' {
			b.WriteByte('"')
			i++
			continue
		}
		inString = !inString
	}
	return b.String()
}

func compactStatement(stmt string) string {
	return strings.ReplaceAll(strings.ReplaceAll(stmt, " ", ""), "\t", "")
}

func isObjectType(typ string) bool {
	typ = strings.ToLower(cleanIdentifier(strings.TrimSpace(typ)))
	if typ == "" {
		return false
	}
	return objectTypes[typ] || strings.HasSuffix(typ, ".application") || strings.HasSuffix(typ, ".workbook") || strings.HasSuffix(typ, ".worksheet") || strings.HasSuffix(typ, ".range") || strings.HasSuffix(typ, ".dictionary")
}

func lastName(name string) string {
	parts := strings.Split(name, ".")
	return parts[len(parts)-1]
}

func referencedTraceHelpers(code string) []string {
	seen := map[string]bool{}
	helpers := make([]string, 0, 2)
	for _, re := range []*regexp.Regexp{traceHelperQualRe, traceHelperCallRe} {
		for _, m := range re.FindAllStringSubmatch(code, -1) {
			if len(m) < 2 {
				continue
			}
			name := m[1]
			key := strings.ToLower(name)
			if seen[key] {
				continue
			}
			seen[key] = true
			helpers = append(helpers, name)
		}
	}
	return helpers
}

func resolveWithInfo(expr string, decls map[string]sourceDeclaration) withInfo {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return withInfo{}
	}
	base := expr
	if idx := strings.Index(base, "("); idx >= 0 {
		base = base[:idx]
	}
	base = lastName(strings.TrimSpace(strings.TrimPrefix(base, "Set ")))
	if decl, ok := decls[strings.ToLower(base)]; ok {
		return withInfo{Target: base, Type: decl.Type}
	}
	return withInfo{}
}

func currentWithInfo(stack []withInfo) (withInfo, bool) {
	for i := len(stack) - 1; i >= 0; i-- {
		if stack[i].Type != "" {
			return stack[i], true
		}
	}
	return withInfo{}, false
}

func invalidMemberRuleFor(typ, member string) (invalidMemberRule, bool) {
	rules, ok := invalidObjectMembers[strings.ToLower(cleanIdentifier(strings.TrimSpace(typ)))]
	if !ok {
		return invalidMemberRule{}, false
	}
	rule, ok := rules[strings.ToLower(cleanIdentifier(strings.TrimSpace(member)))]
	if !ok {
		return invalidMemberRule{}, false
	}
	return rule, true
}

func rangeFindAssignment(stmt string) (string, bool) {
	lower := strings.ToLower(stmt)
	if !strings.Contains(lower, ".find(") && !strings.Contains(lower, ".find ") {
		return "", false
	}
	left, _, ok := strings.Cut(stmt, "=")
	if !ok {
		return "", false
	}
	left = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(left), "Set "))
	fields := strings.FieldsFunc(left, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	if len(fields) == 0 {
		return "", false
	}
	return cleanIdentifier(fields[len(fields)-1]), true
}

func labelName(stmt string) (string, bool) {
	stmt = strings.TrimSpace(stmt)
	if !strings.HasSuffix(stmt, ":") || strings.Contains(stmt, " ") {
		return "", false
	}
	name := cleanIdentifier(strings.TrimSuffix(stmt, ":"))
	return name, name != ""
}

func isCleanupFallthroughLabel(label string) bool {
	switch strings.ToLower(label) {
	case "cleanup", "clean_up", "finally", "done":
		return true
	default:
		return false
	}
}

func errorHandlerFallthroughSuggestion(proc sourceProcedure, label string) string {
	exitStmt := "Exit Sub"
	switch strings.ToLower(proc.Kind) {
	case "function":
		exitStmt = "Exit Function"
	case "property":
		exitStmt = "Exit Property"
	}
	return "Add `" + exitStmt + "` before `" + label + ":`, or rename the label to Cleanup if normal fallthrough is intentional."
}

func terminatesNormalFlow(stmt string) bool {
	lower := strings.ToLower(strings.TrimSpace(stmt))
	return strings.HasPrefix(lower, "exit sub") ||
		strings.HasPrefix(lower, "exit function") ||
		strings.HasPrefix(lower, "exit property") ||
		strings.HasPrefix(lower, "goto ") ||
		lower == "end"
}

func applicationStateName(prop string) string {
	switch strings.ToLower(prop) {
	case "enableevents":
		return "EnableEvents"
	case "displayalerts":
		return "DisplayAlerts"
	case "screenupdating":
		return "ScreenUpdating"
	case "calculation":
		return "Calculation"
	default:
		return prop
	}
}

func parseSimpleCall(stmt string) (string, []string, bool) {
	stmt = strings.TrimSpace(stmt)
	if stmt == "" {
		return "", nil, false
	}
	if strings.HasPrefix(strings.ToLower(stmt), "call ") {
		stmt = strings.TrimSpace(stmt[len("call "):])
	}
	if strings.Contains(stmt, "=") || strings.HasPrefix(strings.ToLower(stmt), "if ") {
		return "", nil, false
	}
	name := ""
	argText := ""
	if idx := strings.Index(stmt, "("); idx >= 0 && strings.HasSuffix(stmt, ")") {
		name = strings.TrimSpace(stmt[:idx])
		argText = strings.TrimSuffix(stmt[idx+1:], ")")
	} else {
		fields := strings.Fields(stmt)
		if len(fields) < 2 {
			return "", nil, false
		}
		name = fields[0]
		argText = strings.TrimSpace(stmt[len(name):])
	}
	name = cleanIdentifier(lastName(name))
	if name == "" {
		return "", nil, false
	}
	return name, splitArgs(argText), true
}

func splitArgs(text string) []string {
	var args []string
	start := 0
	inString := false
	depth := 0
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
		case ',':
			if !inString && depth == 0 {
				args = append(args, strings.TrimSpace(text[start:i]))
				start = i + 1
			}
		}
	}
	if strings.TrimSpace(text[start:]) != "" {
		args = append(args, strings.TrimSpace(text[start:]))
	}
	return args
}

func obviousArgumentMismatch(arg, typ string) bool {
	arg = strings.TrimSpace(arg)
	if arg == "" {
		return false
	}
	lowerType := strings.ToLower(cleanIdentifier(typ))
	isStringLiteral := strings.HasPrefix(arg, `"`) && strings.HasSuffix(arg, `"`)
	isNumericLiteral := true
	for _, r := range arg {
		if (r < '0' || r > '9') && r != '.' && r != '-' {
			isNumericLiteral = false
			break
		}
	}
	if isStringLiteral {
		return lowerType != "string" && lowerType != "variant"
	}
	if isNumericLiteral {
		return lowerType == "string" || isObjectType(typ)
	}
	return false
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

func nodeName(node *tree_sitter.Node, source []byte) string {
	if name := node.ChildByFieldName("name"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
	}
	if name := firstNamedChildKind(node, "identifier"); name != nil {
		return cleanIdentifier(name.Utf8Text(source))
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

func cleanIdentifier(text string) string {
	text = strings.TrimSpace(text)
	text = strings.Trim(text, "[]")
	text = strings.TrimRight(text, "$%&#@^!")
	return text
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

func sortFindings(findings []Finding) {
	sort.SliceStable(findings, func(i, j int) bool {
		a, b := findings[i], findings[j]
		if a.File != b.File {
			return a.File < b.File
		}
		if a.Line != b.Line {
			return a.Line < b.Line
		}
		if a.Column != b.Column {
			return a.Column < b.Column
		}
		return a.Code < b.Code
	})
}

func applyInlineSuppressions(findings []Finding, directives []suppression.Directive) ([]Finding, []map[string]any) {
	diagnostics := make([]suppression.Diagnostic, 0, len(findings))
	for _, finding := range findings {
		diagnostics = append(diagnostics, suppression.Diagnostic{
			Code: finding.Code,
			File: finding.File,
			Line: finding.Line,
		})
	}
	suppressed, warnings := suppression.Apply(diagnostics, directives, suppression.FamilyAnalyze)
	if len(suppressed) == 0 {
		return findings, warnings
	}
	filtered := make([]Finding, 0, len(findings))
	for i, finding := range findings {
		if suppressed[i] {
			continue
		}
		filtered = append(filtered, finding)
	}
	return filtered, warnings
}

func strconvItoa(v int) string {
	const digits = "0123456789"
	if v == 0 {
		return "0"
	}
	var buf [20]byte
	i := len(buf)
	for v > 0 {
		i--
		buf[i] = digits[v%10]
		v /= 10
	}
	return string(buf[i:])
}
