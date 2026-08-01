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
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/suppression"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Finding struct {
	Code         string   `json:"code"`
	Severity     string   `json:"severity"`
	File         string   `json:"file"`
	Module       string   `json:"module,omitempty"`
	Procedure    string   `json:"procedure,omitempty"`
	Line         int      `json:"line"`
	Column       int      `json:"column,omitempty"`
	EndLine      int      `json:"-"`
	EndColumn    int      `json:"-"`
	ScopeEndLine int      `json:"scope_end_line,omitempty"`
	Message      string   `json:"message"`
	Reason       string   `json:"reason"`
	Suggestion   string   `json:"suggestion"`
	NearbyCode   []string `json:"nearby_code,omitempty"`
}

type Result struct {
	Findings []Finding
	Warnings []map[string]any
}

type Analyzer struct {
	RootDir            string
	Config             config.Config
	PathFilter         func(string) bool
	typeDB             *vbadb.DB
	errorGuardAliases  map[string]bool
	errorValueWrappers map[string]bool
}

var (
	declRe                        = regexp.MustCompile(`(?i)^\s*(?:dim|private|public|static)\s+(.+)$`)
	assignRe                      = regexp.MustCompile(`(?i)^\s*(?:let\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	setAssignRe                   = regexp.MustCompile(`(?i)^\s*set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=`)
	callAssignRe                  = regexp.MustCompile(`(?i)^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([A-Za-z_][A-Za-z0-9_.]*)\s*(?:\(|$)`)
	withRe                        = regexp.MustCompile(`(?i)^\s*with\s+(.+)$`)
	endWithRe                     = regexp.MustCompile(`(?i)^\s*end\s+with\b`)
	withMemberRe                  = regexp.MustCompile(`(?i)^\s*\.([A-Za-z_][A-Za-z0-9_]*)\b`)
	memberRe                      = regexp.MustCompile(`(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)\b`)
	traceHelperCallRe             = regexp.MustCompile(`(?i)^\s*(?:call\s+)?(XlflowLog|XlflowSetTraceFile)\b`)
	traceHelperQualRe             = regexp.MustCompile(`(?i)\bXlflowTrace\s*\.\s*(XlflowLog|XlflowSetTraceFile)\b`)
	unqualifiedExcelRe            = regexp.MustCompile(`(?i)(^|[^A-Za-z0-9_.$])(?:Application\s*\.\s*)?\b(Range|Cells|Rows|Columns)\b\s*(?:\(|\.)`)
	activeExcelRe                 = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_.$])(?:Application\s*\.\s*)?\b(ActiveWorkbook|ActiveSheet|ActiveCell|Selection)\b`)
	unqualifiedSheetCollectionRe  = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_.$])(?:Application\s*\.\s*)?\b(Worksheets|Sheets)\b\s*\(`)
	positionalExcelCollectionRe   = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_.$])(?:Application\s*\.\s*)?\b(Workbooks|Windows)\b\s*\(\s*([0-9]+)\s*\)`)
	workbooksOpenRe               = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_.$])((?:Application\s*\.\s*)?\bWorkbooks\s*\.\s*Open\b)`)
	workbooksOpenBranchBoundaryRe = regexp.MustCompile(`(?i):|\bthen\b|\belse\b`)
	setAssignmentPrefixRe         = regexp.MustCompile(`(?i)^\s*set\s+.+?\s*=\s*$`)
	thisWorkbookRe                = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_.$])\bThisWorkbook\b`)
	redimPreserveRe               = regexp.MustCompile(`(?i)^\s*redim\s+preserve\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)`)
	forEachDirectRe               = regexp.MustCompile(`(?i)^\s*for\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+([A-Za-z_][A-Za-z0-9_]*)\s*$`)
	forStartRe                    = regexp.MustCompile(`(?i)^\s*for\b`)
	nextRe                        = regexp.MustCompile(`(?i)^\s*next\b`)
	dictionaryCreateRe            = regexp.MustCompile(`(?i)^\s*createobject\s*\(\s*"scripting\.dictionary"\s*\)\s*$`)
	dictionaryNewRe               = regexp.MustCompile(`(?i)^\s*new\s+scripting\.dictionary\s*$`)
	errProbeReferenceRe           = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_])err\s*\.\s*(?:number|clear)\b`)
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
	functionReturns    map[string]string
	procedures         map[string]procedureSignature
	worksheetCodenames map[string]string
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
	Path       string
	Lines      []string
	Module     string
	ModuleKind string
	Source     []byte
	Root       *tree_sitter.Node
	IR         procedureir.DocumentIR
	CFG        vbacfg.Document
	Parsed     *vbaast.ParsedDocument
}

type sourceProcedure struct {
	Kind       string
	Name       string
	ReturnType string
	StartLine  int
	EndLine    int
	Params     []parameterInfo
	Statements []procedureir.Statement
	Calls      []procedureir.CallSite
	Accesses   []procedureir.VariableAccess
	Graph      *vbacfg.Graph
	Effects    *effects.ProcedureSummary
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
	parsedFiles := make([]parsedFile, 0, len(files))
	for _, file := range files {
		source, err := os.ReadFile(file)
		if err != nil {
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		parsed, err := vbaast.ParseDocument(file, source)
		if err != nil {
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		moduleKind := ""
		if sourceFile, included, classifyErr := symbols.SourceFileForPath(a.RootDir, a.Config, file); classifyErr != nil {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, classifyErr
		} else if included {
			moduleKind = sourceFile.ModuleKind
		}
		ir, err := procedureir.BuildParsed(procedureir.BuildOptions{
			RootDir:    a.RootDir,
			Path:       file,
			ModuleKind: moduleKind,
		}, parsed)
		if err != nil {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		if ir.Parse.HasError || ir.Parse.HasMissing {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, fmt.Errorf("parse %s: VBA parser reported errors or missing nodes", file)
		}
		controlFlow := vbacfg.BuildDocument(ir)
		parsedFiles = append(parsedFiles, parsedFile{
			Path:       file,
			Lines:      normalizedSourceLines(string(source)),
			Module:     strings.TrimSuffix(filepath.Base(file), filepath.Ext(file)),
			ModuleKind: moduleKind,
			Source:     source,
			IR:         ir,
			CFG:        controlFlow,
			Parsed:     parsed,
		})
	}
	defer closeParsedFiles(parsedFiles)

	projectEffects := buildProjectEffects(parsedFiles)
	ctx := a.buildContext(parsedFiles)
	analysis := a
	if a.Config.Analyze.DetectStatefulExcelCallArguments || a.Config.Analyze.DetectExcelAPIFailureContracts {
		typeDB, err := vbadb.LoadBuiltin()
		if err != nil {
			return Result{}, err
		}
		analysis.typeDB = typeDB
	}
	var findings []Finding
	if analysis.Config.Analyze.DetectExcelAPIFailureContracts {
		analysis.errorGuardAliases = projectIsErrorGuardAliases(parsedFiles)
		analysis.errorValueWrappers = projectErrorValueWrappers(parsedFiles)
	}
	for _, file := range parsedFiles {
		if err := file.Parsed.Read(func(view vbaast.ParsedView) error {
			file.Root = view.Root
			findings = append(findings, analysis.analyzeParsedFile(file, ctx, projectEffects)...)
			findings = append(findings, analysis.statefulExcelCallArgumentFindings(file)...)
			findings = append(findings, analysis.excelAPIFailureContractFindings(file)...)
			findings = append(findings, analysis.errorValueWrapperFindings(file)...)
			return nil
		}); err != nil {
			return Result{}, err
		}
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

func (a Analyzer) statefulExcelCallArgumentFindings(file parsedFile) []Finding {
	if !a.Config.Analyze.DetectStatefulExcelCallArguments || a.typeDB == nil {
		return nil
	}
	diagnostics := (intel.Analyzer{RootDir: a.RootDir, Config: a.Config, DB: a.typeDB}).StatefulExcelCallArgumentDiagnostics(intel.Document{
		Path: file.Path, Source: string(file.Source),
	})
	if len(diagnostics) == 0 {
		return nil
	}
	procedures := sourceProceduresFromIR(file.IR, file.CFG)
	out := make([]Finding, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		line := diagnostic.Range.Start.Line + 1
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		for _, candidate := range procedures {
			if line >= candidate.StartLine && line <= candidate.EndLine {
				proc = candidate
				break
			}
		}
		finding := a.simpleFinding(
			file,
			proc,
			line,
			diagnostic.Code,
			diagnostic.Severity,
			diagnostic.Message,
			"Excel saves selected Range.Find and Range.Replace settings; later calls that omit them can inherit settings changed by the Find or Replace dialog or a previous macro call.",
			"Pass every stateful argument explicitly, or suppress VBA215 on a call that intentionally inherits Excel's saved settings.",
		)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out
}

func buildProjectEffects(files []parsedFile) effects.ProjectSummary {
	resolverSymbols := make([]procedureir.ResolverSymbol, 0)
	for _, file := range files {
		for _, proc := range file.IR.Procedures {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name:       proc.Symbol.Name,
				Module:     file.IR.ModuleName,
				ModuleKind: file.IR.ModuleKind,
				Kind:       string(proc.Symbol.Kind),
				Visibility: proc.Symbol.Visibility,
				File:       file.IR.Path,
				Line:       proc.Symbol.DeclarationRange.StartLine,
			})
		}
	}
	resolver := procedureir.NewResolver(resolverSymbols)
	documents := make([]effects.Document, 0, len(files))
	for i := range files {
		files[i].IR = procedureir.Resolve(files[i].IR, resolver)
		documents = append(documents, effects.Document{IR: files[i].IR, CFG: files[i].CFG})
	}
	return effects.Build(documents)
}

func closeParsedFiles(files []parsedFile) {
	for _, file := range files {
		if file.Parsed != nil {
			file.Parsed.Close()
		}
	}
}

func SourceNonShortCircuitObjectGuardFindings(rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	if !cfg.Analyze.DetectNonShortCircuitObjectGuard {
		return nil, nil
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return nil, err
	}
	defer doc.Close()
	return SourceNonShortCircuitObjectGuardFindingsParsed(rootDir, cfg, doc)
}

// SourceNonShortCircuitObjectGuardFindingsParsed analyzes a caller-owned
// parsed VBA document without closing it or retaining tree-sitter nodes.
func SourceNonShortCircuitObjectGuardFindingsParsed(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: rootDir}, doc)
	if err != nil {
		return nil, err
	}
	var findings []Finding
	err = doc.Read(func(view vbaast.ParsedView) error {
		file := parsedFile{
			Path:   view.Path,
			Lines:  normalizedSourceLines(string(view.Source)),
			Module: strings.TrimSuffix(filepath.Base(view.Path), filepath.Ext(view.Path)),
			Source: view.Source,
			Root:   view.Root,
			IR:     ir,
		}
		analyzer := Analyzer{RootDir: rootDir, Config: cfg}
		procedures := sourceProceduresFromIR(ir)
		for _, proc := range procedures {
			findings = append(findings, analyzer.nonShortCircuitObjectGuardFindings(file, proc)...)
		}
		if len(procedures) == 0 {
			findings = append(findings, analyzer.nonShortCircuitObjectGuardFindings(file, sourceProcedure{StartLine: 1, EndLine: len(file.Lines)})...)
		}
		sortFindings(findings)
		directives, _ := suppression.DirectivesForSource(rootDir, view.Path, string(view.Source))
		findings, _ = applyInlineSuppressions(findings, directives)
		return nil
	})
	return findings, err
}

func SourceRealtimeFindings(rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	if !sourceRealtimeAnalysisEnabled(cfg.Analyze) {
		return nil, nil
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return nil, err
	}
	defer doc.Close()
	return SourceRealtimeFindingsParsed(rootDir, cfg, doc)
}

// SourceRealtimeFindingsParsed runs real-time source analysis against a
// caller-owned parsed VBA document. It does not close doc or retain nodes.
func SourceRealtimeFindingsParsed(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	ir, err := procedureir.BuildParsed(procedureir.BuildOptions{RootDir: rootDir}, doc)
	if err != nil {
		return nil, err
	}
	return SourceRealtimeFindingsParsedIR(rootDir, cfg, doc, ir)
}

// SourceRealtimeFindingsParsedIR runs real-time source analysis with a
// caller-supplied procedure IR built from doc. It lets immutable LSP snapshots
// reuse their cached IR without reparsing or rewalking procedure syntax.
func SourceRealtimeFindingsParsedIR(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFG(rootDir, cfg, doc, ir, vbacfg.BuildDocument(ir))
}

// SourceRealtimeFindingsParsedIRCFG runs real-time source analysis with
// caller-supplied procedure IR and control-flow graphs. Immutable LSP
// snapshots use this entry point to reuse both cached analysis layers.
func SourceRealtimeFindingsParsedIRCFG(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGWithTypeDB(rootDir, cfg, doc, ir, controlFlow, nil)
}

// SourceRealtimeFindingsParsedIRCFGWithTypeDB runs real-time source analysis
// with an optional caller-owned type database. LSP callers pass their loaded
// database; standalone callers load the built-in database only when VBA215 is
// enabled.
func SourceRealtimeFindingsParsedIRCFGWithTypeDB(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document, typeDB *vbadb.DB) ([]Finding, error) {
	if !sourceRealtimeAnalysisEnabled(cfg.Analyze) {
		return nil, nil
	}
	if (cfg.Analyze.DetectStatefulExcelCallArguments || cfg.Analyze.DetectExcelAPIFailureContracts) && typeDB == nil {
		var err error
		typeDB, err = vbadb.LoadBuiltin()
		if err != nil {
			return nil, err
		}
	}
	var findings []Finding
	err := doc.Read(func(view vbaast.ParsedView) error {
		file := parsedFile{
			Path:   view.Path,
			Lines:  normalizedSourceLines(string(view.Source)),
			Module: strings.TrimSuffix(filepath.Base(view.Path), filepath.Ext(view.Path)),
			Source: view.Source,
			Root:   view.Root,
			IR:     ir,
			CFG:    controlFlow,
		}
		analyzer := Analyzer{RootDir: rootDir, Config: cfg, typeDB: typeDB}
		worksheetCodenames := realtimeWorksheetCodenames(rootDir, cfg.Src.Workbook, view.Path)
		procedures := sourceProceduresFromIR(ir, controlFlow)
		moduleDecls := moduleDeclarations(file.Lines, procedures)
		if len(procedures) == 0 {
			procedures = []sourceProcedure{{StartLine: 1, EndLine: len(file.Lines)}}
		}
		for _, proc := range procedures {
			findings = append(findings, analyzer.sourceRealtimeProcedureFindings(file, proc, moduleDecls, worksheetCodenames)...)
		}
		findings = append(findings, analyzer.statefulExcelCallArgumentFindings(file)...)
		findings = append(findings, analyzer.excelAPIFailureContractFindings(file)...)
		findings = append(findings, analyzer.errorValueWrapperFindings(file)...)
		findings = realtimeFindings(findings)
		sortFindings(findings)
		directives, _ := suppression.DirectivesForSource(rootDir, view.Path, string(view.Source))
		findings, _ = applyInlineSuppressions(findings, directives)
		return nil
	})
	return findings, err
}

var sourceRealtimeRuleIDs = []string{"VBA201", "VBA204", "VBA208", "VBA209", "VBA212", "VBA213", "VBA215", "VBA216", "VBA217", "VBA218"}

func sourceRealtimeAnalysisEnabled(cfg config.AnalyzeConfig) bool {
	for _, rule := range staticrules.ByFamily(staticrules.FamilyAnalyze) {
		if !rule.Realtime {
			continue
		}
		if enabled, configurable := config.AnalyzeRuleEnabled(cfg, rule.ID); configurable {
			if enabled {
				return true
			}
			continue
		}
		if rule.DefaultEnabled {
			return true
		}
	}
	return false
}

func realtimeFindings(findings []Finding) []Finding {
	out := findings[:0]
	for _, finding := range findings {
		rule, ok := staticrules.Lookup(finding.Code)
		if ok && rule.Realtime {
			out = append(out, finding)
		}
	}
	return out
}

func (a Analyzer) sourceRealtimeProcedureFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, worksheetCodenames map[string]string) []Finding {
	decls := cloneDeclarations(moduleDecls)
	for key, decl := range procedureDeclarations(file.Lines, proc) {
		decls[key] = decl
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Line: proc.StartLine, Object: isObjectType(param.Type), Parameter: true}
	}
	findAssignments := map[string]int{}
	guardedFinds := map[string]bool{}
	worksheetRoots := newWorksheetRootTracker(worksheetCodenames)
	var findings []Finding
	if a.Config.Analyze.DetectNonShortCircuitObjectGuard {
		findings = append(findings, a.nonShortCircuitObjectGuardFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectDictionaryIterationValueUsage {
		findings = append(findings, a.dictionaryIterationValueUsageFindings(file, proc, moduleDecls)...)
	}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		worksheetStmt, worksheetStatementStart := worksheetLogicalStatement(file.Lines, i, proc.EndLine-1)
		if stmt == "" {
			continue
		}
		if endWithRe.MatchString(stmt) {
			worksheetRoots.popWith()
			continue
		}
		if m := withRe.FindStringSubmatch(stmt); len(m) > 0 {
			if worksheetStatementStart {
				if a.Config.Analyze.DetectWorksheetRootMismatch || a.Config.Analyze.DetectUnstableLastRowPatterns {
					findings = append(findings, a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)...)
				}
				if rootWith := withRe.FindStringSubmatch(worksheetStmt); len(rootWith) > 0 {
					worksheetRoots.pushWith(rootWith[1])
				} else {
					worksheetRoots.pushWith(m[1])
				}
			}
			continue
		}
		if worksheetStatementStart {
			if a.Config.Analyze.DetectWorksheetRootMismatch || a.Config.Analyze.DetectUnstableLastRowPatterns {
				findings = append(findings, a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)...)
			}
			worksheetRoots.observeSetAssignment(worksheetStmt)
		}
		if setAssignRe.MatchString(stmt) {
			if name, ok := rangeFindAssignment(stmt); ok {
				findAssignments[strings.ToLower(name)] = lineNo
			}
			continue
		}
		if a.Config.Analyze.DetectRangeFindNothingCheck {
			findings = append(findings, a.rangeFindFindings(file, proc, lineNo, stmt, findAssignments, guardedFinds)...)
		}
		if a.Config.Analyze.DetectRedimPreserveDimension {
			findings = append(findings, a.redimPreserveFindings(file, proc, lineNo, stmt)...)
		}
		if a.Config.Analyze.DetectObjectArrayComparison {
			findings = append(findings, a.objectArrayComparisonFindings(file, proc, lineNo, stmt, decls)...)
		}
	}
	if a.Config.Analyze.DetectErrorHandlerFallthrough {
		findings = append(findings, a.errorHandlerFallthroughFindings(file, proc)...)
	}
	return findings
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
	ctx := analysisContext{
		functionReturns:    map[string]string{},
		procedures:         map[string]procedureSignature{},
		worksheetCodenames: map[string]string{},
	}
	workbookRoot := filepath.Clean(filepath.Join(a.RootDir, a.Config.Src.Workbook))
	for _, file := range files {
		if rel, err := filepath.Rel(workbookRoot, file.Path); err == nil && rel != "" && !strings.HasPrefix(rel, "..") && !strings.EqualFold(file.Module, "ThisWorkbook") {
			ctx.worksheetCodenames[strings.ToLower(file.Module)] = file.Module
		}
		for _, proc := range sourceProceduresFromIR(file.IR) {
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

func (a Analyzer) analyzeParsedFile(file parsedFile, ctx analysisContext, projectEffects effects.ProjectSummary) []Finding {
	reportedMissingHelpers := map[string]bool{}
	var findings []Finding
	procedures := sourceProceduresFromIR(file.IR, file.CFG)
	for i := range procedures {
		if i >= len(file.IR.Procedures) {
			break
		}
		symbol := file.IR.Procedures[i].Symbol
		id := effects.ProcedureIdentity{
			File: file.IR.Path, Module: file.IR.ModuleName, ModuleKind: file.IR.ModuleKind, Name: symbol.Name,
			QualifiedName: symbol.QualifiedName, Kind: symbol.Kind,
			Visibility: symbol.Visibility, DeclarationLine: symbol.DeclarationRange.StartLine,
		}
		if summary, ok := projectEffects.Lookup(id); ok {
			procedures[i].Effects = &summary
		}
	}
	moduleDecls := moduleDeclarations(file.Lines, procedures)
	for _, proc := range procedures {
		findings = append(findings, a.analyzeProcedure(file, proc, moduleDecls, ctx, projectEffects, reportedMissingHelpers)...)
	}
	if len(procedures) == 0 {
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		findings = append(findings, a.analyzeProcedure(file, proc, moduleDecls, ctx, projectEffects, reportedMissingHelpers)...)
	}
	return findings
}

func (a Analyzer) analyzeProcedure(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, projectEffects effects.ProjectSummary, reportedMissingHelpers map[string]bool) []Finding {
	decls := cloneDeclarations(moduleDecls)
	for key, decl := range procedureDeclarations(file.Lines, proc) {
		decls[key] = decl
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Line: proc.StartLine, Object: isObjectType(param.Type), Parameter: true}
	}
	withStack := make([]withInfo, 0)
	initialized := initialObjectState(decls)
	maybeInitializedByCall := map[string]bool{}
	findAssignments := map[string]int{}
	guardedFinds := map[string]bool{}
	functionAssigned := false
	worksheetRoots := newWorksheetRootTracker(ctx.worksheetCodenames)
	var findings []Finding
	if a.Config.Analyze.DetectNonShortCircuitObjectGuard {
		findings = append(findings, a.nonShortCircuitObjectGuardFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectDictionaryIterationValueUsage {
		findings = append(findings, a.dictionaryIterationValueUsageFindings(file, proc, moduleDecls)...)
	}

	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		worksheetStmt, worksheetStatementStart := worksheetLogicalStatement(file.Lines, i, proc.EndLine-1)
		if stmt == "" {
			continue
		}
		lower := strings.ToLower(stmt)

		if endWithRe.MatchString(stmt) {
			if len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
			worksheetRoots.popWith()
			continue
		}
		if worksheetStatementStart && a.Config.Analyze.ForbidUnqualifiedExcelObjects {
			findings = append(findings, a.unqualifiedExcelFindings(file, proc, lineNo, normalizedCodeLine(worksheetStmt))...)
		}
		if m := withRe.FindStringSubmatch(stmt); len(m) > 0 {
			withStack = append(withStack, resolveWithInfo(m[1], decls))
			if worksheetStatementStart {
				if a.Config.Analyze.DetectWorksheetRootMismatch || a.Config.Analyze.DetectUnstableLastRowPatterns {
					findings = append(findings, a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)...)
				}
				if rootWith := withRe.FindStringSubmatch(worksheetStmt); len(rootWith) > 0 {
					worksheetRoots.pushWith(rootWith[1])
				} else {
					worksheetRoots.pushWith(m[1])
				}
			}
			continue
		}
		if worksheetStatementStart {
			if a.Config.Analyze.DetectWorksheetRootMismatch || a.Config.Analyze.DetectUnstableLastRowPatterns {
				findings = append(findings, a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)...)
			}
			worksheetRoots.observeSetAssignment(worksheetStmt)
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
		findings = append(findings, a.applicationStateFindings(file, proc, projectEffects)...)
	}
	if a.Config.Analyze.DetectErrorHandlerFallthrough {
		findings = append(findings, a.errorHandlerFallthroughFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectLeakedOnErrorResumeNextScopes {
		findings = append(findings, a.leakedOnErrorResumeNextFindings(file, proc)...)
	}
	if a.Config.Analyze.DetectFunctionReturnPath && proc.Kind == "Function" && proc.Name != "" && !functionAssigned {
		findings = append(findings, a.simpleFinding(file, proc, proc.StartLine, "VBA210", "warning", proc.Name+" may exit without assigning its return value.", "Functions return the default value when no assignment to the function name is reached.", "Assign "+proc.Name+" on every successful return path, or make the default return explicit."))
	}
	return findings
}

func sourceProceduresFromIR(document procedureir.DocumentIR, controlFlow ...vbacfg.Document) []sourceProcedure {
	procedures := make([]sourceProcedure, 0, len(document.Procedures))
	for procedureIndex, procedure := range document.Procedures {
		kind := "Sub"
		switch procedure.Symbol.Kind {
		case procedureir.ProcedureFunction:
			kind = "Function"
		case procedureir.ProcedureProperty, procedureir.ProcedurePropertyGet,
			procedureir.ProcedurePropertyLet, procedureir.ProcedurePropertySet:
			kind = "Property"
		}
		params := make([]parameterInfo, len(procedure.Symbol.Parameters))
		for i, parameter := range procedure.Symbol.Parameters {
			params[i] = parameterInfo{
				Name: parameter.Name, Type: parameter.Type, Passing: parameter.Passing,
			}
		}
		source := sourceProcedure{
			Kind:       kind,
			Name:       procedure.Symbol.Name,
			ReturnType: procedure.Symbol.ReturnType,
			StartLine:  procedure.Symbol.DeclarationRange.StartLine,
			EndLine:    procedure.Symbol.DeclarationRange.EndLine,
			Params:     params,
			Statements: append([]procedureir.Statement(nil), procedure.Statements...),
			Calls:      append([]procedureir.CallSite(nil), procedure.Calls...),
			Accesses:   append([]procedureir.VariableAccess(nil), procedure.Accesses...),
		}
		if len(controlFlow) > 0 && procedureIndex < len(controlFlow[0].Graphs) {
			graph := controlFlow[0].Graphs[procedureIndex]
			source.Graph = &graph
		}
		procedures = append(procedures, source)
	}
	return procedures
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
		name := m[1]
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", name+" creates an active Excel object dependency.", name+" depends on the user's current Excel UI state and can target a different object during automation.", "Pass an explicit Workbook, Worksheet, or Range argument instead."))
	}
	for _, m := range unqualifiedSheetCollectionRe.FindAllStringSubmatch(stmt, -1) {
		name := m[1]
		suggestion := "Use ThisWorkbook." + name + "(...) or select " + name + " from an explicit Workbook argument."
		if addInStandardModule(a.Config, file) {
			suggestion = "Select " + name + " from an explicit caller Workbook argument."
		}
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Unqualified "+name+" access depends on the active workbook.", "The unqualified "+name+" collection is resolved from Excel's active workbook at runtime.", suggestion))
	}
	for _, m := range positionalExcelCollectionRe.FindAllStringSubmatch(stmt, -1) {
		name := m[1]
		index := m[2]
		root := name + "(" + index + ")"
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", root+" depends on Excel collection ordering.", root+" can select a different object when workbook or window order changes.", "Select the target by name or receive an explicit "+strings.TrimSuffix(name, "s")+" argument."))
	}
	for _, open := range workbooksOpenRe.FindAllStringSubmatchIndex(stmt, -1) {
		if !capturedWorkbooksOpen(stmt, open[2]) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Workbooks.Open result is not captured.", "An uncaptured Workbooks.Open result forces later code to depend on active workbook state.", "Capture the opened workbook: Set wb = Workbooks.Open(...)."))
		}
	}
	if addInStandardModule(a.Config, file) && thisWorkbookRe.MatchString(stmt) {
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "ThisWorkbook in an add-in targets the add-in workbook.", "Inside an add-in standard module, ThisWorkbook is the add-in rather than the caller workbook.", "Receive the caller workbook as an explicit Workbook argument."))
	}
	return findings
}

func addInStandardModule(cfg config.Config, file parsedFile) bool {
	return strings.EqualFold(filepath.Ext(cfg.Excel.Path), ".xlam") && file.ModuleKind == "standard"
}

func capturedWorkbooksOpen(stmt string, openStart int) bool {
	if openStart < 0 || openStart > len(stmt) {
		return false
	}
	branchStart := 0
	for _, boundary := range workbooksOpenBranchBoundaryRe.FindAllStringIndex(stmt[:openStart], -1) {
		branchStart = boundary[1]
	}
	return setAssignmentPrefixRe.MatchString(stmt[branchStart:openStart])
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

type dictionaryIterationLoop struct {
	item     string
	dict     string
	reported bool
}

// dictionaryIterationValueUsageFindings reports only direct dictionary
// iteration whose control variable is subsequently used as an object. VBA
// iterates Dictionary keys, so ordinary key usage must remain silent.
func (a Analyzer) dictionaryIterationValueUsageFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) []Finding {
	decls := cloneDeclarations(moduleDecls)
	for key, decl := range procedureDeclarations(file.Lines, proc) {
		decls[key] = decl
	}
	for _, param := range proc.Params {
		decls[strings.ToLower(param.Name)] = sourceDeclaration{Name: param.Name, Type: param.Type, Line: proc.StartLine, Object: isObjectType(param.Type), Parameter: true}
	}

	declaredDictionaries := map[string]bool{}
	for key, decl := range decls {
		if isDictionaryType(decl.Type) {
			declaredDictionaries[key] = true
		}
	}
	inferredDictionaries := map[string]bool{}
	var loops []*dictionaryIterationLoop
	var findings []Finding

	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		rawStmt := strings.Join(strings.Fields(gui.StripComment(file.Lines[i])), " ")
		if stmt == "" {
			continue
		}

		if nextRe.MatchString(stmt) {
			if len(loops) > 0 {
				loops = loops[:len(loops)-1]
			}
			continue
		}
		if m := forEachDirectRe.FindStringSubmatch(stmt); len(m) > 0 {
			dictKey := strings.ToLower(m[2])
			loop := &dictionaryIterationLoop{}
			if declaredDictionaries[dictKey] || inferredDictionaries[dictKey] {
				loop.item = m[1]
				loop.dict = m[2]
			}
			loops = append(loops, loop)
			continue
		}
		if forStartRe.MatchString(stmt) {
			loops = append(loops, &dictionaryIterationLoop{})
			continue
		}

		if m := setAssignRe.FindStringSubmatch(stmt); len(m) > 0 {
			target := strings.ToLower(m[1])
			for _, loop := range loops {
				if loop.item != "" && strings.EqualFold(loop.item, m[1]) {
					loop.item = ""
				}
			}
			rhs := strings.TrimSpace(rawStmt[strings.Index(rawStmt, "=")+1:])
			if dictionaryCreateRe.MatchString(rhs) || dictionaryNewRe.MatchString(rhs) {
				inferredDictionaries[target] = true
			} else {
				delete(inferredDictionaries, target)
			}
		}

		for _, loop := range loops {
			if loop.item == "" || loop.reported || !dictionaryIterationValueUse(stmt, loop.item, decls) {
				continue
			}
			findings = append(findings, a.simpleFinding(
				file,
				proc,
				lineNo,
				"VBA213",
				"warning",
				loop.item+" is iterated directly from "+loop.dict+" but is used as an object/value.",
				"Direct Scripting.Dictionary iteration yields keys, not items or values.",
				"Iterate "+loop.dict+".Items for values, or retrieve the value with "+loop.dict+"("+loop.item+") before using it as an object.",
			))
			loop.reported = true
		}
	}
	return findings
}

func isDictionaryType(typ string) bool {
	switch strings.ToLower(cleanIdentifier(typ)) {
	case "dictionary", "scripting.dictionary":
		return true
	default:
		return false
	}
}

func dictionaryIterationValueUse(stmt, item string, decls map[string]sourceDeclaration) bool {
	if match := withRe.FindStringSubmatch(stmt); len(match) > 1 && strings.EqualFold(cleanIdentifier(match[1]), item) {
		return true
	}
	itemKey := strings.ToLower(item)
	for _, match := range memberRe.FindAllStringSubmatch(stmt, -1) {
		if len(match) > 1 && strings.EqualFold(match[1], item) {
			return true
		}
	}
	match := setAssignRe.FindStringSubmatch(stmt)
	if len(match) == 0 {
		return false
	}
	target, ok := decls[strings.ToLower(match[1])]
	if !ok || !target.Object {
		return false
	}
	rhs := strings.TrimSpace(stmt[strings.Index(stmt, "=")+1:])
	return strings.EqualFold(cleanIdentifier(rhs), itemKey)
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
		if decl.Array && identifierComparedAsOperand(lower, key, proc) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA209", "warning", decl.Name+" appears to be compared as a scalar value.", "VBA arrays cannot be compared directly to scalar values.", "Compare explicit elements or bounds instead of the array variable itself."))
		}
	}
	return findings
}

func identifierComparedAsOperand(stmt, name string, proc sourceProcedure) bool {
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
		if opLen == 1 && stmt[i] == '=' && isFunctionReturnAssignment(stmt, i, proc) {
			i += opLen - 1
			continue
		}
		if operandEndsWithBareIdentifier(left, name) || operandStartsWithBareIdentifier(right, name) {
			return true
		}
		i += opLen - 1
	}
	return false
}

func isFunctionReturnAssignment(stmt string, operatorIndex int, proc sourceProcedure) bool {
	match := assignRe.FindStringSubmatchIndex(stmt)
	if len(match) < 4 || match[1]-1 != operatorIndex || proc.Kind != "Function" || proc.Name == "" {
		return false
	}
	return strings.EqualFold(stmt[match[2]:match[3]], proc.Name)
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

func (a Analyzer) applicationStateFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	if proc.Graph == nil {
		return nil
	}
	byID := make(map[int]procedureir.Statement, len(proc.Statements))
	for _, statement := range proc.Statements {
		byID[statement.ID] = statement
	}
	var findings []Finding
	for _, property := range applicationStateProperties() {
		unsafe := applicationStateExitWitnesses(proc, property.Key, byID)
		if len(unsafe) == 0 || hasPairedApplicationRestoreProcedure(proc, property.Key, project) {
			continue
		}
		for _, statement := range proc.Statements {
			witness, found := unsafe[statement.ID]
			if !found {
				continue
			}
			assigned, _, ok := applicationPropertyAssignment(statement, byID)
			if !ok || assigned != property.Key {
				continue
			}
			findings = append(findings, a.simpleFinding(
				file, proc, statement.Range.StartLine, "VBA203", "warning",
				"Application."+property.Name+" can reach "+witness.Kind+" without restoring its previous value.",
				"The changed Application."+property.Name+" value can leave this procedure through "+witness.description()+".",
				"Save the previous Application."+property.Name+" value and restore it in a cleanup path.",
			))
		}
	}
	return findings
}

type applicationStateProperty struct {
	Key  string
	Name string
}

func applicationStateProperties() []applicationStateProperty {
	return []applicationStateProperty{
		{Key: "screenupdating", Name: "ScreenUpdating"},
		{Key: "enableevents", Name: "EnableEvents"},
		{Key: "displayalerts", Name: "DisplayAlerts"},
		{Key: "calculation", Name: "Calculation"},
		{Key: "statusbar", Name: "StatusBar"},
		{Key: "cursor", Name: "Cursor"},
		{Key: "interactive", Name: "Interactive"},
		{Key: "asktoupdatelinks", Name: "AskToUpdateLinks"},
		{Key: "automationsecurity", Name: "AutomationSecurity"},
		{Key: "cutcopymode", Name: "CutCopyMode"},
	}
}

type applicationStateSnapshot struct {
	Dirty   map[int]bool
	Unknown bool
}

type applicationStateFlow struct {
	Dirty          map[int]bool
	Saved          map[string]applicationStateSnapshot
	ViaExceptional bool
}

type applicationStateExitWitness struct {
	Kind string
	Line int
}

func (w applicationStateExitWitness) description() string {
	if w.Line > 0 {
		return w.Kind + " at line " + strconvItoa(w.Line)
	}
	return w.Kind
}

// applicationStateExitWitnesses performs a may analysis. A state-changing
// assignment is reported whenever one possible path reaches an exit while the
// prior state is still dirty. State-changing assignments take effect only on
// their normal successors; exceptional successors observe the input state
// because VBA can fault before an assignment completes. A proven saved-value
// restoration is the cleanup boundary itself, so it clears state on its own
// exceptional transition as well.
func applicationStateExitWitnesses(proc sourceProcedure, property string, byID map[int]procedureir.Statement) map[int]applicationStateExitWitness {
	graph := proc.Graph
	in := map[vbacfg.BlockID]applicationStateFlow{graph.Entry: newApplicationStateFlow()}
	queued := map[vbacfg.BlockID]bool{graph.Entry: true}
	queue := []vbacfg.BlockID{graph.Entry}
	witnesses := map[int]applicationStateExitWitness{}

	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		queued[blockID] = false
		state := in[blockID]
		block, ok := applicationStateBlock(*graph, blockID)
		if !ok {
			continue
		}
		for _, edge := range graph.Edges {
			if edge.From != blockID {
				continue
			}
			out := cloneApplicationStateFlow(state)
			out.ViaExceptional = out.ViaExceptional || edge.Class == vbacfg.EdgeExceptional
			if block.Statement != nil && (edge.Class == vbacfg.EdgeNormal || applicationStateSavedRestore(proc, state, *block.Statement, property, byID)) {
				out = applyApplicationStateStatement(proc, out, *block.Statement, property, byID)
			}
			if applicationStateExitKind(*graph, edge.To) != "" {
				kind := applicationStateExitKind(*graph, edge.To)
				if out.ViaExceptional && kind == "normal exit" {
					kind = "error-handler path to normal exit"
				}
				for origin := range out.Dirty {
					if existing, exists := witnesses[origin]; !exists || applicationStateWitnessRank(kind) > applicationStateWitnessRank(existing.Kind) {
						witnesses[origin] = applicationStateExitWitness{Kind: kind, Line: edge.Range.StartLine}
					}
				}
			}
			merged, changed := mergeApplicationStateFlow(in[edge.To], out)
			if !changed {
				continue
			}
			in[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
		}
	}
	return witnesses
}

func applicationStateWitnessRank(kind string) int {
	switch kind {
	case "error-handler path to normal exit":
		return 4
	case "error exit":
		return 3
	case "unknown exit":
		return 2
	case "termination exit":
		return 1
	default:
		return 0
	}
}

func applicationStateBlock(graph vbacfg.Graph, id vbacfg.BlockID) (vbacfg.Block, bool) {
	if id <= 0 || int(id) > len(graph.Blocks) {
		return vbacfg.Block{}, false
	}
	return graph.Blocks[int(id)-1], true
}

func applicationStateExitKind(graph vbacfg.Graph, id vbacfg.BlockID) string {
	switch id {
	case graph.NormalExit:
		return "normal exit"
	case graph.ExceptionalExit:
		return "error exit"
	case graph.TerminationExit:
		return "termination exit"
	case graph.UnknownExit:
		return "unknown exit"
	default:
		return ""
	}
}

func newApplicationStateFlow() applicationStateFlow {
	return applicationStateFlow{Dirty: map[int]bool{}, Saved: map[string]applicationStateSnapshot{}}
}

func cloneApplicationStateFlow(in applicationStateFlow) applicationStateFlow {
	out := newApplicationStateFlow()
	for origin := range in.Dirty {
		out.Dirty[origin] = true
	}
	for key, snapshot := range in.Saved {
		out.Saved[key] = cloneApplicationStateSnapshot(snapshot)
	}
	out.ViaExceptional = in.ViaExceptional
	return out
}

func cloneApplicationStateSnapshot(in applicationStateSnapshot) applicationStateSnapshot {
	out := applicationStateSnapshot{Dirty: map[int]bool{}, Unknown: in.Unknown}
	for origin := range in.Dirty {
		out.Dirty[origin] = true
	}
	return out
}

func mergeApplicationStateFlow(current, incoming applicationStateFlow) (applicationStateFlow, bool) {
	if current.Dirty == nil {
		return cloneApplicationStateFlow(incoming), true
	}
	next := cloneApplicationStateFlow(current)
	changed := false
	for origin := range incoming.Dirty {
		if !next.Dirty[origin] {
			next.Dirty[origin] = true
			changed = true
		}
	}
	if incoming.ViaExceptional && !next.ViaExceptional {
		next.ViaExceptional = true
		changed = true
	}
	keys := map[string]bool{}
	for key := range current.Saved {
		keys[key] = true
	}
	for key := range incoming.Saved {
		keys[key] = true
	}
	for key := range keys {
		left, leftOK := current.Saved[key]
		right, rightOK := incoming.Saved[key]
		if !leftOK {
			left = applicationStateSnapshot{Dirty: map[int]bool{}, Unknown: true}
		}
		if !rightOK {
			right = applicationStateSnapshot{Dirty: map[int]bool{}, Unknown: true}
		}
		merged := cloneApplicationStateSnapshot(left)
		merged.Unknown = left.Unknown || right.Unknown
		for origin := range right.Dirty {
			merged.Dirty[origin] = true
		}
		if !applicationStateSnapshotEqual(current.Saved[key], merged) {
			next.Saved[key] = merged
			changed = true
		}
	}
	return next, changed
}

func applicationStateSnapshotEqual(a, b applicationStateSnapshot) bool {
	if a.Unknown != b.Unknown || len(a.Dirty) != len(b.Dirty) {
		return false
	}
	for origin := range a.Dirty {
		if !b.Dirty[origin] {
			return false
		}
	}
	return true
}

func applyApplicationStateStatement(proc sourceProcedure, state applicationStateFlow, statement procedureir.Statement, property string, byID map[int]procedureir.Statement) applicationStateFlow {
	if statement.Recovered || statement.Kind != procedureir.StatementAssignment || statement.Target == nil {
		return state
	}
	if assignedProperty, value, isPropertyWrite := applicationPropertyAssignment(statement, byID); isPropertyWrite && assignedProperty == property {
		if applicationStateSavedRestore(proc, state, statement, property, byID) {
			variable, _ := applicationStateVariable(proc, statement.ID, value, procedureir.AccessRead)
			state.Dirty = cloneApplicationStateSnapshot(state.Saved[variable]).Dirty
			return state
		}
		state.Dirty[statement.ID] = true
		return state
	}
	variable, ok := applicationStateVariable(proc, statement.ID, statement.Target.Text, procedureir.AccessWrite)
	if !ok {
		return state
	}
	if isApplicationPropertyReference(statement.Value, statement, byID, property) {
		state.Saved[variable] = applicationStateSnapshot{Dirty: cloneApplicationStateFlow(state).Dirty}
		return state
	}
	if source, ok := applicationStateVariable(proc, statement.ID, expressionText(statement.Value), procedureir.AccessRead); ok {
		if saved, exists := state.Saved[source]; exists {
			state.Saved[variable] = cloneApplicationStateSnapshot(saved)
			return state
		}
	}
	state.Saved[variable] = applicationStateSnapshot{Dirty: map[int]bool{}, Unknown: true}
	return state
}

func applicationStateSavedRestore(proc sourceProcedure, state applicationStateFlow, statement procedureir.Statement, property string, byID map[int]procedureir.Statement) bool {
	assignedProperty, value, isPropertyWrite := applicationPropertyAssignment(statement, byID)
	if !isPropertyWrite || assignedProperty != property {
		return false
	}
	variable, ok := applicationStateVariable(proc, statement.ID, value, procedureir.AccessRead)
	if !ok {
		return false
	}
	saved, exists := state.Saved[variable]
	return exists && !saved.Unknown
}

func expressionText(expr *procedureir.Expression) string {
	if expr == nil {
		return ""
	}
	return expr.Text
}

func applicationStateVariable(proc sourceProcedure, statementID int, expression string, mode procedureir.AccessMode) (string, bool) {
	name := cleanIdentifier(strings.TrimSpace(expression))
	if name == "" || strings.ContainsAny(name, ".() ") {
		return "", false
	}
	for _, access := range proc.Accesses {
		if access.StatementID != statementID || !strings.EqualFold(access.Name, name) {
			continue
		}
		if mode == procedureir.AccessWrite && access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite {
			continue
		}
		if mode == procedureir.AccessRead && access.Mode != procedureir.AccessRead && access.Mode != procedureir.AccessReadWrite {
			continue
		}
		if access.Scope == procedureir.ScopeUnresolved {
			return "", false
		}
		return string(access.Scope) + ":" + strings.ToLower(name), true
	}
	return "", false
}

func applicationPropertyAssignment(statement procedureir.Statement, byID map[int]procedureir.Statement) (string, string, bool) {
	if statement.Target == nil || statement.Value == nil {
		return "", "", false
	}
	property, ok := applicationPropertyTarget(statement.Target.Text, statement, byID)
	if !ok {
		return "", "", false
	}
	return property, statement.Value.Text, true
}

func isApplicationPropertyReference(expr *procedureir.Expression, statement procedureir.Statement, byID map[int]procedureir.Statement, property string) bool {
	if expr == nil {
		return false
	}
	got, ok := applicationPropertyTarget(expr.Text, statement, byID)
	return ok && got == property
}

func applicationPropertyTarget(expression string, statement procedureir.Statement, byID map[int]procedureir.Statement) (string, bool) {
	compact := strings.ToLower(compactStatement(expression))
	for _, property := range applicationStateProperties() {
		if compact == "application."+property.Key {
			return property.Key, true
		}
		if compact == "."+property.Key && statementWithinApplicationWith(statement, byID) {
			return property.Key, true
		}
	}
	return "", false
}

func statementWithinApplicationWith(statement procedureir.Statement, byID map[int]procedureir.Statement) bool {
	for parentID := statement.ParentID; parentID != 0; {
		parent, ok := byID[parentID]
		if !ok {
			return false
		}
		if parent.Kind == procedureir.StatementWith {
			return strings.EqualFold(compactStatement(parent.Text), "withapplication")
		}
		parentID = parent.ParentID
	}
	return false
}

func hasPairedApplicationRestoreProcedure(proc sourceProcedure, prop string, project effects.ProjectSummary) bool {
	if proc.Name == "" || proc.Effects == nil {
		return false
	}
	lowerName := strings.ToLower(proc.Name)
	if strings.HasPrefix(lowerName, "push") && hasApplicationStateEffect(proc.Effects.Direct, effects.ChangesApplicationState, prop) {
		suffix := strings.TrimPrefix(lowerName, "push")
		pairNames := map[string]bool{"pop" + suffix: true, "restore" + suffix: true}
		sameModule := make([]effects.ProcedureSummary, 0, 1)
		projectVisible := make([]effects.ProcedureSummary, 0, 1)
		for _, summary := range project.All() {
			if !pairNames[strings.ToLower(summary.Identity.Name)] {
				continue
			}
			if strings.EqualFold(summary.Identity.Module, proc.Effects.Identity.Module) {
				sameModule = append(sameModule, summary)
				continue
			}
			if isProjectVisibleProcedure(summary.Identity) {
				projectVisible = append(projectVisible, summary)
			}
		}
		if len(sameModule) > 0 {
			for _, candidate := range sameModule {
				if hasApplicationStateEffect(candidate.Direct, effects.RestoresApplicationState, prop) ||
					hasApplicationStateEffect(candidate.Propagated, effects.RestoresApplicationState, prop) {
					return true
				}
			}
			return false
		}
		if len(projectVisible) == 1 {
			return hasApplicationStateEffect(projectVisible[0].Direct, effects.RestoresApplicationState, prop) ||
				hasApplicationStateEffect(projectVisible[0].Propagated, effects.RestoresApplicationState, prop)
		}
	}

	// Keep the established helper convention quiet on the restore side too. A
	// restoration helper can be the paired Pop/Restore procedure itself or a
	// uniquely resolved leaf called by that procedure.
	if !hasApplicationStateEffect(proc.Effects.Direct, effects.RestoresApplicationState, prop) {
		return false
	}
	for _, candidate := range project.All() {
		if !hasApplicationStateEffectFrom(candidate.Direct, effects.RestoresApplicationState, prop, proc.Effects.Identity) &&
			!hasApplicationStateEffectFrom(candidate.Propagated, effects.RestoresApplicationState, prop, proc.Effects.Identity) {
			continue
		}
		if hasMatchingPushProcedure(candidate, prop, project) {
			return true
		}
	}
	return false
}

func hasApplicationStateEffect(evidence []effects.Evidence, kind effects.EffectKind, prop string) bool {
	target := "application." + strings.ToLower(prop)
	for _, item := range evidence {
		if item.Effect == kind && strings.EqualFold(item.Target, target) {
			return true
		}
	}
	return false
}

func hasApplicationStateEffectFrom(evidence []effects.Evidence, kind effects.EffectKind, prop string, origin effects.ProcedureIdentity) bool {
	target := "application." + strings.ToLower(prop)
	for _, item := range evidence {
		if item.Effect == kind && strings.EqualFold(item.Target, target) && item.Origin.Key() == origin.Key() {
			return true
		}
	}
	return false
}

func hasMatchingPushProcedure(candidate effects.ProcedureSummary, prop string, project effects.ProjectSummary) bool {
	name := strings.ToLower(candidate.Identity.Name)
	var suffix string
	switch {
	case strings.HasPrefix(name, "pop"):
		suffix = strings.TrimPrefix(name, "pop")
	case strings.HasPrefix(name, "restore"):
		suffix = strings.TrimPrefix(name, "restore")
	default:
		return false
	}
	matches := 0
	for _, push := range project.All() {
		if !strings.EqualFold(push.Identity.Name, "push"+suffix) || !hasApplicationStateEffect(push.Direct, effects.ChangesApplicationState, prop) {
			continue
		}
		if strings.EqualFold(push.Identity.Module, candidate.Identity.Module) || isProjectVisibleProcedure(push.Identity) && isProjectVisibleProcedure(candidate.Identity) {
			return true
		}
		matches++
	}
	// This mirrors the existing cross-module Push/Pop exception: one visible
	// Pop helper may establish a paired restore even when its Push counterpart
	// is private in a different standard module.
	return matches == 1 && isProjectVisibleProcedure(candidate.Identity)
}

func isProjectVisibleProcedure(identity effects.ProcedureIdentity) bool {
	return (identity.ModuleKind == "" || strings.EqualFold(identity.ModuleKind, "standard")) &&
		!strings.EqualFold(strings.TrimSpace(identity.Visibility), "private")
}

func (a Analyzer) errorHandlerFallthroughFindings(file parsedFile, proc sourceProcedure) []Finding {
	handlerLabels := map[string]bool{}
	for _, statement := range proc.Statements {
		if statement.Kind != procedureir.StatementOnError {
			continue
		}
		label := cleanIdentifier(statement.Label)
		if label != "" && !strings.EqualFold(label, "0") {
			handlerLabels[strings.ToLower(label)] = true
		}
	}
	if len(handlerLabels) == 0 {
		return nil
	}
	var findings []Finding
	if proc.Graph != nil {
		reachable := map[vbacfg.BlockID]bool{}
		for _, blockID := range proc.Graph.Reachable(vbacfg.EdgeFilter{NormalOnly: true}) {
			reachable[blockID] = true
		}
		for _, statement := range proc.Statements {
			if statement.Kind != procedureir.StatementLabel {
				continue
			}
			label := cleanIdentifier(statement.Label)
			if !handlerLabels[strings.ToLower(label)] || isCleanupFallthroughLabel(label) {
				continue
			}
			block, ok := proc.Graph.BlockForStatement(statement.ID)
			if !ok || !reachable[block.ID] {
				continue
			}
			implicitEntry := false
			for _, edge := range proc.Graph.Edges {
				if edge.To == block.ID && edge.Class == vbacfg.EdgeNormal && reachable[edge.From] &&
					edge.Kind != vbacfg.EdgeGoto && edge.Kind != vbacfg.EdgeUnknown {
					implicitEntry = true
					break
				}
			}
			if !implicitEntry {
				continue
			}
			lineNo := statement.Range.StartLine
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA204", "warning", "Normal execution can fall through into error handler "+label+".", "Without Exit Sub, Exit Function, or Exit Property before the handler label, successful execution can run error handling code.", errorHandlerFallthroughSuggestion(proc, label)))
		}
		return findings
	}

	lastCodeByParent := map[int]string{}
	for _, statement := range proc.Statements {
		if statement.Kind == procedureir.StatementLabel {
			label := cleanIdentifier(statement.Label)
			if handlerLabels[strings.ToLower(label)] &&
				!isCleanupFallthroughLabel(label) && !terminatesNormalFlow(lastCodeByParent[statement.ParentID]) {
				lineNo := statement.Range.StartLine
				findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA204", "warning", "Normal execution can fall through into error handler "+label+".", "Without Exit Sub, Exit Function, or Exit Property before the handler label, successful execution can run error handling code.", errorHandlerFallthroughSuggestion(proc, label)))
			}
		}
		lastCodeByParent[statement.ParentID] = statement.Text
	}
	return findings
}

type resumeNextScopeFlag uint8

const (
	resumeNextScopeMultipleOperations resumeNextScopeFlag = 1 << iota
	resumeNextScopeControlFlow
	resumeNextScopeCall
	resumeNextScopeProjectCall
	resumeNextScopeUnrestoredExit
)

type resumeNextScopeState struct {
	block      vbacfg.BlockID
	operations uint8
	flags      resumeNextScopeFlag
}

type resumeNextScopeOutcome struct {
	startLine int
	endLine   int
	flags     resumeNextScopeFlag
}

// leakedOnErrorResumeNextFindings follows each reachable Resume Next scope until
// it is explicitly replaced or leaves the procedure. The small state domain is
// deliberately saturated at two operations: the rule only needs to distinguish
// a single compatibility probe from a wider protected region.
func (a Analyzer) leakedOnErrorResumeNextFindings(file parsedFile, proc sourceProcedure) []Finding {
	if proc.Graph == nil {
		return nil
	}
	reachable := map[vbacfg.BlockID]bool{}
	for _, id := range proc.Graph.Reachable(vbacfg.EdgeFilter{}) {
		reachable[id] = true
	}
	outcomes := map[string]resumeNextScopeOutcome{}
	for _, statement := range proc.Statements {
		if !isOnErrorResumeNext(statement) {
			continue
		}
		start, ok := proc.Graph.BlockForStatement(statement.ID)
		if !ok || !reachable[start.ID] {
			continue
		}
		for _, outcome := range resumeNextScopeOutcomes(proc, start.ID) {
			if outcome.flags == 0 {
				continue
			}
			outcome.startLine = statement.Range.StartLine
			key := strings.Join([]string{strconvItoa(outcome.startLine), strconvItoa(outcome.endLine)}, ":")
			if existing, found := outcomes[key]; found {
				existing.flags |= outcome.flags
				outcomes[key] = existing
			} else {
				outcomes[key] = outcome
			}
		}
	}
	if len(outcomes) == 0 {
		return nil
	}
	keys := make([]string, 0, len(outcomes))
	for key := range outcomes {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	findings := make([]Finding, 0, len(keys))
	for _, key := range keys {
		outcome := outcomes[key]
		severity := "warning"
		if outcome.flags&resumeNextScopeProjectCall != 0 {
			severity = "error"
		}
		finding := a.simpleFinding(
			file, proc, outcome.startLine, "VBA214", severity,
			"On Error Resume Next starting at line "+strconvItoa(outcome.startLine)+" remains active through line "+strconvItoa(outcome.endLine)+".",
			resumeNextScopeReason(outcome.flags),
			"Limit `On Error Resume Next` to one compatibility probe, inspect and clear `Err` when needed, then restore error handling with `On Error GoTo 0` before calls, branches, or exits.",
		)
		finding.ScopeEndLine = outcome.endLine
		findings = append(findings, finding)
	}
	return findings
}

func resumeNextScopeOutcomes(proc sourceProcedure, start vbacfg.BlockID) []resumeNextScopeOutcome {
	graph := proc.Graph
	queue := make([]resumeNextScopeState, 0)
	var outcomes []resumeNextScopeOutcome
	for _, edge := range graph.Edges {
		if edge.From != start || edge.Class != vbacfg.EdgeNormal {
			continue
		}
		if isResumeNextScopeExit(*graph, edge.To) {
			outcomes = append(outcomes, resumeNextScopeOutcome{
				endLine: resumeNextScopeEndLine(proc, *graph, edge),
				flags:   resumeNextScopeUnrestoredExit,
			})
			continue
		}
		queue = append(queue, resumeNextScopeState{block: edge.To})
	}
	seen := map[string]bool{}
	for len(queue) > 0 {
		state := queue[0]
		queue = queue[1:]
		key := strings.Join([]string{strconvItoa(int(state.block)), strconvItoa(int(state.operations)), strconvItoa(int(state.flags))}, ":")
		if seen[key] {
			continue
		}
		seen[key] = true
		block, ok := resumeNextScopeBlock(*graph, state.block)
		if !ok || block.Statement == nil {
			continue
		}
		statement := *block.Statement
		if restoresErrorHandling(statement) {
			if state.flags != 0 {
				outcomes = append(outcomes, resumeNextScopeOutcome{endLine: statement.Range.StartLine, flags: state.flags})
			}
			continue
		}
		state = applyResumeNextScopeStatement(proc, state, statement)
		for _, edge := range graph.Edges {
			if edge.From != state.block {
				continue
			}
			if !resumeNextScopeFollowsEdge(*graph, edge) {
				continue
			}
			if isResumeNextScopeExit(*graph, edge.To) {
				terminal := state
				terminal.flags |= resumeNextScopeUnrestoredExit
				outcomes = append(outcomes, resumeNextScopeOutcome{
					endLine: resumeNextScopeEndLine(proc, *graph, edge),
					flags:   terminal.flags,
				})
				continue
			}
			queue = append(queue, resumeNextScopeState{block: edge.To, operations: state.operations, flags: state.flags})
		}
	}
	return outcomes
}

func isOnErrorResumeNext(statement procedureir.Statement) bool {
	return statement.Kind == procedureir.StatementOnError && statement.Control != nil &&
		statement.Control.Transfer == procedureir.TransferOnErrorResumeNext
}

func restoresErrorHandling(statement procedureir.Statement) bool {
	if statement.Kind != procedureir.StatementOnError || statement.Control == nil {
		return false
	}
	return statement.Control.Transfer == procedureir.TransferOnErrorDisable ||
		statement.Control.Transfer == procedureir.TransferOnErrorGoto
}

func applyResumeNextScopeStatement(proc sourceProcedure, state resumeNextScopeState, statement procedureir.Statement) resumeNextScopeState {
	if isOnErrorResumeNext(statement) || resumeNextScopeErrProbeStatement(statement) ||
		statement.Kind == procedureir.StatementDeclaration || statement.Kind == procedureir.StatementLabel {
		return state
	}
	call, projectCall := resumeNextScopeCallRisk(proc.Calls, statement.ID)
	if call {
		state.flags |= resumeNextScopeCall
	}
	if projectCall {
		state.flags |= resumeNextScopeProjectCall
	}
	if resumeNextScopeControlStatement(statement) {
		state.flags |= resumeNextScopeControlFlow
		return state
	}
	if state.operations < 2 {
		state.operations++
	}
	if state.operations >= 2 {
		state.flags |= resumeNextScopeMultipleOperations
	}
	return state
}

func resumeNextScopeErrProbeStatement(statement procedureir.Statement) bool {
	code := maskStringLiterals(gui.StripComment(statement.Text))
	return errProbeReferenceRe.MatchString(code)
}

func resumeNextScopeControlStatement(statement procedureir.Statement) bool {
	switch statement.Kind {
	case procedureir.StatementIf, procedureir.StatementElseIf, procedureir.StatementElse,
		procedureir.StatementSelect, procedureir.StatementCase, procedureir.StatementFor,
		procedureir.StatementForEach, procedureir.StatementDo, procedureir.StatementWhile,
		procedureir.StatementWith, procedureir.StatementGoTo, procedureir.StatementResume:
		return true
	default:
		return false
	}
}

func resumeNextScopeCallRisk(calls []procedureir.CallSite, statementID int) (call bool, projectCall bool) {
	for _, candidate := range calls {
		if candidate.StatementID != statementID {
			continue
		}
		switch candidate.Resolution.Status {
		case procedureir.ResolutionMatched:
			if len(candidate.Resolution.Candidates) == 1 {
				return true, true
			}
		case procedureir.ResolutionAmbiguous, procedureir.ResolutionUnresolved, procedureir.ResolutionNotAttempted:
			call = true
		}
	}
	return call, projectCall
}

func isResumeNextScopeExit(graph vbacfg.Graph, id vbacfg.BlockID) bool {
	return id == graph.NormalExit || id == graph.ExceptionalExit || id == graph.TerminationExit || id == graph.UnknownExit
}

// A graph can merge paths with different active error modes. Only an
// exceptional EdgeError that mirrors a normal successor represents Resume Next
// continuation for this scope; handler and disabled-mode error edges belong to
// the other merged mode and must not be followed.
func resumeNextScopeFollowsEdge(graph vbacfg.Graph, edge vbacfg.Edge) bool {
	if edge.Class != vbacfg.EdgeExceptional {
		return true
	}
	if edge.Kind != vbacfg.EdgeError {
		return false
	}
	for _, normal := range graph.Edges {
		if normal.From == edge.From && normal.To == edge.To && normal.Class == vbacfg.EdgeNormal {
			return true
		}
	}
	return false
}

func resumeNextScopeBlock(graph vbacfg.Graph, id vbacfg.BlockID) (vbacfg.Block, bool) {
	if id <= 0 || int(id) > len(graph.Blocks) {
		return vbacfg.Block{}, false
	}
	return graph.Blocks[int(id)-1], true
}

func resumeNextScopeEndLine(proc sourceProcedure, graph vbacfg.Graph, edge vbacfg.Edge) int {
	if edge.To == graph.NormalExit && edge.Kind != vbacfg.EdgeProcedureExit {
		return proc.EndLine
	}
	if edge.Range.StartLine > 0 {
		return edge.Range.StartLine
	}
	return proc.EndLine
}

func resumeNextScopeReason(flags resumeNextScopeFlag) string {
	switch {
	case flags&resumeNextScopeProjectCall != 0:
		return "A resolved project-local procedure call executes while `On Error Resume Next` is active."
	case flags&resumeNextScopeCall != 0:
		return "A procedure call can execute while `On Error Resume Next` is active."
	case flags&resumeNextScopeUnrestoredExit != 0:
		return "The procedure can exit before error handling is restored."
	case flags&resumeNextScopeMultipleOperations != 0:
		return "More than one executable operation is protected by `On Error Resume Next`."
	default:
		return "Control flow other than an Err.Number check executes while `On Error Resume Next` is active."
	}
}

func BlockingFindings(findings []Finding) []Finding {
	blocking := make([]Finding, 0)
	for _, finding := range findings {
		metadata, ok := staticrules.Lookup(finding.Code)
		if ok && metadata.PreflightBlocking {
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

func (a Analyzer) nonShortCircuitObjectGuardFindings(file parsedFile, proc sourceProcedure) []Finding {
	seen := map[string]bool{}
	var findings []Finding
	var visit func(*tree_sitter.Node)
	visit = func(node *tree_sitter.Node) {
		if node == nil {
			return
		}
		r := vbaast.NodeRange(node)
		if r.StartLine >= proc.StartLine && (proc.EndLine == 0 || r.StartLine <= proc.EndLine) &&
			isBooleanBinaryExpression(node) && hasTopLevelAndOrOperator(node, file.Source) {
			guards := map[string]string{}
			accesses := map[string]string{}
			collectNothingGuards(node, file.Source, guards)
			collectDirectMemberAccesses(node, file.Source, accesses)
			for key, name := range guards {
				if _, ok := accesses[key]; !ok {
					continue
				}
				dedupeKey := strconvItoa(r.StartLine) + ":" + key
				if seen[dedupeKey] {
					continue
				}
				seen[dedupeKey] = true
				findings = append(findings, a.simpleFinding(
					file,
					proc,
					r.StartLine,
					"VBA212",
					"warning",
					name+" is guarded against Nothing and dereferenced in the same non-short-circuit boolean expression.",
					"VBA And/Or expressions do not short-circuit, so the member access can still run when the object is Nothing and raise runtime error 91.",
					"Split the Nothing guard and the member access into separate If statements.",
				))
			}
			return
		}
		for i := uint(0); i < node.NamedChildCount(); i++ {
			visit(node.NamedChild(i))
		}
	}
	visit(file.Root)
	return findings
}

func isBooleanBinaryExpression(node *tree_sitter.Node) bool {
	if node == nil {
		return false
	}
	switch node.Kind() {
	case "condition_binary_expression", "binary_expression":
		return true
	default:
		return false
	}
}

func hasTopLevelAndOrOperator(node *tree_sitter.Node, source []byte) bool {
	if node == nil {
		return false
	}
	if node.Kind() == "condition_binary_expression" {
		return true
	}
	text := strings.ToLower(node.Utf8Text(source))
	return hasWord(text, "And") || hasWord(text, "Or")
}

func collectNothingGuards(node *tree_sitter.Node, source []byte, guards map[string]string) {
	if node == nil {
		return
	}
	if node.Kind() == "comparison_expression" {
		if name, ok := nothingGuardIdentifier(node, source); ok {
			guards[strings.ToLower(name)] = name
		}
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		collectNothingGuards(node.NamedChild(i), source, guards)
	}
}

func nothingGuardIdentifier(node *tree_sitter.Node, source []byte) (string, bool) {
	left := node.ChildByFieldName("left")
	right := node.ChildByFieldName("right")
	operator := node.ChildByFieldName("operator")
	if left == nil || right == nil || operator == nil || !strings.EqualFold(strings.TrimSpace(operator.Utf8Text(source)), "Is") {
		return "", false
	}
	if left.Kind() == "identifier" && right.Kind() == "nothing_literal" {
		name := cleanIdentifier(left.Utf8Text(source))
		return name, name != ""
	}
	if right.Kind() == "identifier" && left.Kind() == "nothing_literal" {
		name := cleanIdentifier(right.Utf8Text(source))
		return name, name != ""
	}
	return "", false
}

func collectDirectMemberAccesses(node *tree_sitter.Node, source []byte, accesses map[string]string) {
	if node == nil {
		return
	}
	if node.Kind() == "qualified_member_expression" {
		if name, ok := directMemberReceiverIdentifier(node, source); ok {
			accesses[strings.ToLower(name)] = name
		}
	}
	for i := uint(0); i < node.NamedChildCount(); i++ {
		collectDirectMemberAccesses(node.NamedChild(i), source, accesses)
	}
}

func directMemberReceiverIdentifier(node *tree_sitter.Node, source []byte) (string, bool) {
	receiver := node.ChildByFieldName("receiver")
	if receiver == nil || receiver.Kind() != "identifier" {
		return "", false
	}
	name := cleanIdentifier(receiver.Utf8Text(source))
	return name, name != ""
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

// worksheetLogicalStatement joins only explicit continuation lines for the
// worksheet-root rules. Other rules keep their established physical-line or
// AST handling, while these rules need the complete member chain to compare
// roots accurately. The caller reports on the first physical line.
func worksheetLogicalStatement(lines []string, index, last int) (string, bool) {
	if index < 0 || index >= len(lines) || (index > 0 && vbaLineContinues(lines[index-1])) {
		return "", false
	}
	if last >= len(lines) {
		last = len(lines) - 1
	}
	parts := make([]string, 0, 1)
	for current := index; current <= last; current++ {
		line := rawWorksheetCodeLine(lines[current])
		continues := vbaLineContinues(lines[current])
		if continues {
			line = strings.TrimSpace(strings.TrimSuffix(strings.TrimSpace(line), "_"))
		}
		if line != "" {
			parts = append(parts, line)
		}
		if !continues {
			return strings.Join(parts, " "), true
		}
	}
	return strings.Join(parts, " "), true
}

func rawWorksheetCodeLine(line string) string {
	return strings.TrimSpace(gui.StripComment(line))
}

func vbaLineContinues(line string) bool {
	line = strings.TrimSpace(gui.StripComment(line))
	if len(line) < 2 || line[len(line)-1] != '_' {
		return false
	}
	return line[len(line)-2] == ' ' || line[len(line)-2] == '\t'
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
