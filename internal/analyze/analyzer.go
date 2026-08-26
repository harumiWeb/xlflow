package analyze

import (
	"context"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"unicode"

	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/gui"
	"github.com/harumiWeb/xlflow/internal/lint"
	staticpreflight "github.com/harumiWeb/xlflow/internal/staticanalysis/preflight"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/suppression"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

type Finding struct {
	Code         string            `json:"code"`
	Severity     string            `json:"severity"`
	File         string            `json:"file"`
	Module       string            `json:"module,omitempty"`
	Procedure    string            `json:"procedure,omitempty"`
	Line         int               `json:"line"`
	Column       int               `json:"column,omitempty"`
	EndLine      int               `json:"-"`
	EndColumn    int               `json:"-"`
	ScopeEndLine int               `json:"scope_end_line,omitempty"`
	Message      string            `json:"message"`
	Reason       string            `json:"reason"`
	Suggestion   string            `json:"suggestion"`
	NearbyCode   []string          `json:"nearby_code,omitempty"`
	CallCycle    *CallCycleContext `json:"call_cycle,omitempty"`
	DataFlow     *DataFlowContext  `json:"data_flow,omitempty"`
	// CommandExecution is present on VBA236 findings and augments the generic
	// data-flow context with process-launch role/risk metadata.
	CommandExecution *CommandExecutionContext `json:"command_execution,omitempty"`
	// SQLExecution is present on VBA239 findings and augments the generic
	// data-flow context with SQL API/role/risk metadata.
	SQLExecution *SQLExecutionContext `json:"sql_execution,omitempty"`
	// FileOperation is present on VBA245 findings and carries the operation-
	// specific path safety classification.
	FileOperation *FileOperationContext `json:"file_operation,omitempty"`
	// HTTPSecurity and HTTPReliability are redacted HTTP-specific contexts for
	// VBA246 and VBA247. They never contain header values or complete URLs.
	HTTPSecurity    *HTTPSecurityContext    `json:"http_security,omitempty"`
	HTTPReliability *HTTPReliabilityContext `json:"http_reliability,omitempty"`
	// OpaqueBoolean describes the positional Boolean literals that made a
	// VBA248 call-site finding actionable without embedding parser internals.
	OpaqueBoolean *OpaqueBooleanContext `json:"opaque_boolean,omitempty"`
	// RuntimeError carries the deterministic runtime-failure kind for VBA249.
	// It is deliberately additive so existing finding consumers can continue
	// to consume the common envelope without inferring runtime semantics from
	// severity alone.
	RuntimeError          *RuntimeErrorContext `json:"runtime_error,omitempty"`
	arrayLifecycleFinding bool
	arrayOperationKey     string
	httpOwnedSinks        map[int]bool
	dataFlowSinkStartByte int
}

// ParseError reports that tree-sitter could not produce a complete VBA
// syntax tree for a source file. Callers that need to distinguish parser
// failures from other analysis failures can use errors.As with this type.
type ParseError struct {
	Path       string
	HasError   bool
	HasMissing bool
}

func (e *ParseError) Error() string {
	if e == nil {
		return "VBA parser reported errors or missing nodes"
	}
	return fmt.Sprintf("parse %s: VBA parser reported errors or missing nodes", e.Path)
}

type Result struct {
	Findings          []Finding
	Warnings          []map[string]any
	AnalysisMetrics   any       `json:"analysis_metrics,omitempty"`
	PreflightFindings []Finding `json:"-"`
	// AnalyzedFiles lets regression tests and callers distinguish a successful
	// no-finding run from a run that did not discover any source files. It is
	// intentionally omitted from serialized CLI payloads.
	AnalyzedFiles int `json:"-"`
}

type Analyzer struct {
	RootDir string
	Config  config.Config
	// arrayStrategy is intentionally private and test/benchmark-only.  The
	// zero value selects the production auto compatibility decision; tests may
	// force the indexed or legacy oracle without changing public configuration.
	arrayStrategy              arrayCFGStrategy
	PathFilter                 func(string) bool
	typeDB                     *vbadb.DB
	typeDBResolutionIncomplete bool
	visibleConstants           map[string]bool
	visibleConstantValues      map[string]constexpr.Value
	byRefSymbolIndex           *intel.WorkspaceResolutionView
	errorGuardAliases          map[string]bool
	errorValueWrappers         map[string]bool
	eventSafeProcedures        map[string]bool
	applicationStateLeaks      *applicationStateLeakIndex
	excelLoopAccess            *excelLoopAccessIndex
	excelRootBindings          excelRootBindingIndex
	dictionaryCollection       *dictionaryCollectionIndex
	// analysisWorkerLimit is test-only tuning for the bounded file analysis
	// pool. A zero value derives the limit from GOMAXPROCS and project size.
	analysisWorkerLimit int
	// procedureAnalysisStartHook is test-only synchronization for cancellation
	// coverage. Production callers leave it nil.
	procedureAnalysisStartHook func()
}

// procedureParallelThreshold is deliberately high enough that the ordinary
// file path does not pay for a second scheduler.  The threshold is kept
// internal: callers should tune the existing analysis worker limit rather than
// depending on a new public setting.
const procedureParallelThreshold = 500

// analysisExecutionBudget is shared by file jobs and procedure jobs.  A file
// worker gives its slot back while waiting for its procedure jobs, so nested
// scheduling never multiplies the configured worker count.
type analysisExecutionBudget struct {
	limit         int
	sem           chan struct{}
	procedureJobs chan procedureBatchJob
	procedureWg   sync.WaitGroup
}

type procedureBatchJob struct {
	ctx      context.Context
	run      func(context.Context) ([]Finding, error)
	complete func([]Finding, error)
}

type analysisExecutionBudgetKey struct{}

func newAnalysisExecutionBudget(limit int) *analysisExecutionBudget {
	if limit < 1 {
		limit = 1
	}
	return &analysisExecutionBudget{limit: limit, sem: make(chan struct{}, limit)}
}

func (budget *analysisExecutionBudget) acquire(ctx context.Context) error {
	if budget == nil {
		return ctx.Err()
	}
	select {
	case budget.sem <- struct{}{}:
		return nil
	case <-ctx.Done():
		return ctx.Err()
	}
}

func (budget *analysisExecutionBudget) release() {
	if budget == nil {
		return
	}
	<-budget.sem
}

func (budget *analysisExecutionBudget) startProcedurePool() {
	if budget == nil {
		return
	}
	budget.procedureJobs = make(chan procedureBatchJob)
	budget.procedureWg.Add(budget.limit)
	for worker := 0; worker < budget.limit; worker++ {
		go func() {
			defer budget.procedureWg.Done()
			for job := range budget.procedureJobs {
				if err := budget.acquire(job.ctx); err != nil {
					job.complete(nil, err)
					continue
				}
				findings, err := job.run(job.ctx)
				budget.release()
				job.complete(findings, err)
			}
		}()
	}
}

func (budget *analysisExecutionBudget) submitProcedureBatch(ctx context.Context, job procedureBatchJob) bool {
	if budget == nil || budget.procedureJobs == nil {
		return false
	}
	select {
	case budget.procedureJobs <- job:
		return true
	case <-ctx.Done():
		return false
	}
}

func (budget *analysisExecutionBudget) stopProcedurePool() {
	if budget == nil || budget.procedureJobs == nil {
		return
	}
	close(budget.procedureJobs)
	budget.procedureWg.Wait()
}

func withAnalysisExecutionBudget(ctx context.Context, budget *analysisExecutionBudget) context.Context {
	return context.WithValue(ctx, analysisExecutionBudgetKey{}, budget)
}

func analysisExecutionBudgetFromContext(ctx context.Context) *analysisExecutionBudget {
	if ctx == nil {
		return nil
	}
	budget, _ := ctx.Value(analysisExecutionBudgetKey{}).(*analysisExecutionBudget)
	return budget
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
	forEachDirectRe               = regexp.MustCompile(`(?i)^\s*for\s+each\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+([A-Za-z_][A-Za-z0-9_]*)\s*$`)
	forStartRe                    = regexp.MustCompile(`(?i)^\s*for\b`)
	nextRe                        = regexp.MustCompile(`(?i)^\s*next\b`)
	dictionaryCreateRe            = regexp.MustCompile(`(?i)^\s*createobject\s*\(\s*"scripting\.dictionary"\s*\)\s*$`)
	dictionaryNewRe               = regexp.MustCompile(`(?i)^\s*new\s+scripting\.dictionary\s*$`)
	errProbeReferenceRe           = regexp.MustCompile(`(?i)(?:^|[^A-Za-z0-9_])err\s*\.\s*(?:number|description|clear)\b`)
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
	functionReturns                  map[string]string
	functionShapes                   map[string]procedureir.ValueShapeKind
	functionNamesSeen                map[string]bool
	functionAmbiguous                map[string]bool
	arrayReturns                     map[string]arrayValue
	arrayAllocationGuards            map[string]bool
	arrayByRefAllocations            arrayByRefAllocationSummaries
	arrayByRefConditionalAllocations arrayByRefConditionalAllocations
	arrayByRefLengthAllocations      arrayByRefLengthAllocations
	arrayModuleAllocations           arrayModuleAllocationSummaries
	arrayModuleConfigurations        map[string]arrayModuleConfigurationState
	arrayModuleEntryStates           arrayModuleEntryStates
	arrayPrivateTargets              map[string]sourceProcedure
	arrayParticipants                map[string]bool
	arrayParticipantKeys             map[string]string
	// arrayInterproceduralParticipants excludes complete procedures whose only
	// evidence is an unknown array capability. Those procedures remain in the
	// local participant plan for fail-open diagnostics, but cannot by themselves
	// widen every fixed-point lane in a large module.
	arrayInterproceduralParticipants map[string]bool
	// arrayIgnoreFeatureUnknown is retained for focused participant-set
	// compatibility helpers. Production planning derives both feature-unknown
	// boundaries from one shared graph; complete IR with an unknown array
	// capability remains local unless a real seed reaches that procedure.
	arrayIgnoreFeatureUnknown bool
	arrayStats                *arrayInterproceduralStats
	arrayByRefEntryStates     map[string]map[int]bool
	arrayByRefEntryConditions map[string]map[int]string
	arrayCapabilityIndex      *semanticArrayCapabilityIndex
	procedures                map[string]procedureSignature
	procedureResolver         procedureir.Resolver
	objectAnalysis            *objectAnalysisContext
	worksheetCodenames        map[string]string
	projectEffects            effects.ProjectSummary
	queryRevision             *semanticquery.Revision
}

type arrayInterproceduralStats struct {
	mu                  sync.Mutex
	cfgWalks            uint64
	revisits            uint64
	compactWalks        uint64
	legacyWalks         uint64
	fallbackWalks       uint64
	fallbackEmptyState  uint64
	fallbackIndex       uint64
	fallbackUnsupported uint64
	strategy            arrayCFGStrategy
}

// Array CFG walks can be materialized concurrently while a procedure plan is
// evaluated. Keep the developer-only counters race-free without imposing any
// synchronization on the semantic state itself.
func (s *arrayInterproceduralStats) addCFGWalk() {
	if s == nil {
		return
	}
	s.mu.Lock()
	s.cfgWalks++
	s.mu.Unlock()
}

func (s *arrayInterproceduralStats) addRevisit() {
	if s == nil {
		return
	}
	s.mu.Lock()
	s.revisits++
	s.mu.Unlock()
}

func (s *arrayInterproceduralStats) addCompactWalk() {
	if s == nil {
		return
	}
	s.mu.Lock()
	s.compactWalks++
	s.mu.Unlock()
}

func (s *arrayInterproceduralStats) addLegacyWalk() {
	if s == nil {
		return
	}
	s.mu.Lock()
	s.legacyWalks++
	s.mu.Unlock()
}

func (s *arrayInterproceduralStats) addFallbackReason(reason arrayFallbackReason) {
	if s == nil {
		return
	}
	s.mu.Lock()
	s.fallbackWalks++
	switch reason {
	case arrayFallbackEmptyState:
		s.fallbackEmptyState++
	case arrayFallbackIndex:
		s.fallbackIndex++
	default:
		s.fallbackUnsupported++
	}
	s.mu.Unlock()
}

func (s *arrayInterproceduralStats) snapshot() (cfgWalks, revisits, compactWalks, legacyWalks, fallbackWalks, fallbackEmptyState, fallbackIndex, fallbackUnsupported uint64) {
	if s == nil {
		return
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	return s.cfgWalks, s.revisits, s.compactWalks, s.legacyWalks, s.fallbackWalks, s.fallbackEmptyState, s.fallbackIndex, s.fallbackUnsupported
}

type procedureSignature struct {
	Name       string
	ReturnType string
	Params     readOnlySpan[parameterInfo]
}

// parameterInfo is an internal compatibility alias. Procedure parameters are
// owned by ProcedureIR and are never copied into an analyzer projection.
type parameterInfo = procedureir.Parameter

type parsedFile struct {
	Path       string
	Lines      []string
	Module     string
	ModuleKind string
	Source     []byte
	Root       *tree_sitter.Node
	IR         procedureir.DocumentIR
	CFG        vbacfg.Document
	// Procedures owns the analyzer-facing projection of IR for this file
	// revision. It is materialized once during batch/realtime file setup and
	// reused by all rule stages. Callers must treat the sourceProcedure values
	// and their nested IR/CFG data as read-only; procedures returns an
	// independent outer slice so field updates cannot mutate this cache.
	Procedures []sourceProcedure
	// ModuleDeclarations is the module-scope declaration projection paired with
	// Procedures. It is materialized once during batch/realtime file setup and
	// reused by rule stages that solve array and object state.
	ModuleDeclarations map[string]sourceDeclaration
	// ArrayVariableCatalog is the immutable array-variable projection for
	// module declarations. Procedure-local arrayVariables calls overlay their
	// own declarations on this catalog instead of rescanning every module
	// declaration for every procedure.
	ArrayVariableCatalog map[string]arrayVariable
	// ModuleFacts owns the immutable module declaration and procedure ownership
	// indexes shared by all rule stages for this file revision.
	ModuleFacts                 *moduleAnalysisFacts
	moduleFactsFingerprint      string
	Parsed                      *vbaast.ParsedDocument
	IntelDocument               intel.Document
	RangeValueModuleConstants   map[string]int
	ArrayIntegerModuleConstants map[string]int
	ArrayOptionBase             int
	ArrayOptionBaseSet          bool
	ConstantValues              map[string]constexpr.Value
	DataFlowModuleBindings      map[string]bool
	// semanticQueryFacts is populated once the immutable revision capabilities
	// (effects, resolution and array participant state) are ready.  Procedure
	// query lanes share this pointer so they do not rebuild the same source,
	// dependency, or capability fingerprints for every lane.
	semanticQueryFacts *semanticFileQueryFacts
}

type parsedFileAnalysisResult struct {
	findings  []Finding
	preflight []Finding
	err       error
}

type boundedFileAnalyzer func(context.Context, parsedFile, analysisContext, effects.ProjectSummary, *apiTypeIndex) ([]Finding, []Finding, error)

// batchByRefDiagnostics records the result of the one ByRef analysis pass
// performed for a file revision. The computed bit distinguishes an analyzed
// file with no findings from callers that still need the Intel fallback.
type batchByRefDiagnostics struct {
	computed    bool
	diagnostics []intel.Diagnostic
}

type sourceProcedure struct {
	// Document and IR retain the Go-owned immutable IR for this analysis
	// revision.  IR points at Document.Procedures[Index]; the projection must
	// never retain parser/tree-sitter state.
	Document         *procedureir.DocumentIR
	IR               *procedureir.ProcedureIR
	Index            int
	Kind             string
	ProcedureKind    procedureir.ProcedureKind
	Name             string
	Module           string
	ModuleKind       string
	Visibility       string
	ReturnType       string
	ReturnValueShape procedureir.ValueShapeKind
	StartLine        int
	EndLine          int
	StartByte        int
	EndByte          int
	Params           readOnlySpan[parameterInfo]
	// These are read-only views over IR.  The type is deliberately private; it
	// provides range/len compatibility for the analyzer while exposing no
	// mutable slice through a package boundary.  New consumers should prefer
	// the view/facts accessors.
	Declarations readOnlySpan[procedureir.Declaration]
	Statements   readOnlySpan[procedureir.Statement]
	Expressions  readOnlySpan[procedureir.Expression]
	Calls        readOnlySpan[procedureir.CallSite]
	Accesses     readOnlySpan[procedureir.VariableAccess]
	// Features is the immutable, analyzer-owned applicability summary built
	// with the procedure facts. It contains no parser-owned values.
	Features procedureFeatureSet
	// ArrayParticipant is set after project resolution when this procedure is
	// in the semantic array dependency closure.  A false value is meaningful
	// only for production procedures with complete facts; compatibility
	// projections continue to use their local feature evidence.
	ArrayParticipant      bool
	ArrayParticipantReady bool
	// Plan is the immutable, configuration-specific semantic-domain plan for
	// this procedure revision. PlanReady distinguishes a valid all-skipped plan
	// from projections assembled by focused tests and compatibility helpers.
	Plan      procedureAnalysisPlan
	PlanReady bool
	// Facts is the immutable procedure-local index shared by analysis rules.
	// Its canonical backing storage is IR.  Synthetic focused tests may use
	// the compatibility constructors in procedure_facts.go.
	Facts *procedureAnalysisFacts
	// ModuleFacts is attached only while materializing a plan so applicability
	// can inspect visible sensitive constants without retaining a file-global
	// cache in the plan or result store.
	ModuleFacts *moduleAnalysisFacts
	Graph       *vbacfg.Graph
	Effects     *effects.ProcedureSummary
}

type sourceDeclaration struct {
	Name          string
	Type          string
	Line          int
	Object        bool
	Array         bool
	Fixed         bool
	Dimensions    []arrayDimension
	NewExpression bool
	Parameter     bool
	ParamArray    bool
	Static        bool
}

type withInfo struct {
	Target     string
	Type       string
	Expression string
}

func (a Analyzer) Run() ([]Finding, error) {
	result, err := a.RunResult()
	if err != nil {
		return nil, err
	}
	return result.Findings, nil
}

func (a Analyzer) RunResult() (Result, error) {
	return a.RunResultContext(context.Background())
}

// recordFactBuilds records the construction performed by the current file
// setup. The caller invokes this only after attaching the newly built facts to
// the file, so benchmark counters describe one file/procedure revision rather
// than each rule-family lookup of the shared object.
func recordFactBuilds(ctx context.Context, procedureCount int) {
	recorder := analysisstats.FromContext(ctx)
	if recorder == nil {
		return
	}
	recorder.RecordModuleFactBuild()
	recorder.RecordModuleOptionScan()
	if procedureCount > 0 {
		recorder.RecordProcedureFactBuilds(uint64(procedureCount))
	}
}

// RunContext is the cancellable variant of Run.
func (a Analyzer) RunContext(ctx context.Context) ([]Finding, error) {
	result, err := a.RunResultContext(ctx)
	if err != nil {
		return nil, err
	}
	return result.Findings, nil
}

// RunResultContext is the cancellable variant of RunResult. Cancellation is
// returned explicitly so callers can distinguish it from analysis findings or
// a project-specific timeout policy.
func (a Analyzer) RunResultContext(ctx context.Context) (result Result, err error) {
	if ctx == nil {
		ctx = context.Background()
	}
	ctx, queryContext := withSemanticQueryContext(ctx)
	finishTotal := analysisstats.Measure(ctx, "analyze_total")
	defer func() { finishTotal(len(result.Findings), err) }()
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	finishStage := analysisstats.Measure(ctx, "source_discovery")
	files, err := a.files()
	finishStage(len(files), err)
	if err != nil {
		return Result{}, err
	}
	// Compile-equivalent argument, Set, ByRef, and local-type findings are
	// always enabled because they represent VBE compile rejections and cannot
	// be disabled by the legacy VBA206 runtime-safety setting.
	needsByRefAnalysis := true
	needsTypedExcelAnalysis := a.Config.Analyze.DetectRangeFindNothingCheck || a.Config.Analyze.DetectStatefulExcelCallArguments || a.Config.Analyze.DetectExcelAPIFailureContracts || needsByRefAnalysis || a.Config.Analyze.DetectExcelCellAccessInLoops || a.Config.Analyze.DetectLoopInvariantExcelObjectResolution || a.Config.Analyze.DetectExpensiveFullRangeOperations || a.Config.Analyze.DetectValue2PerformanceOpportunities
	needsDataFlowInputs := dataFlowInputsEnabled(a.Config.Analyze)
	needsTypeDB := needsTypedExcelAnalysis || a.Config.Analyze.DetectPublicAPITypeSafety || needsDataFlowInputs
	parsedFiles := make([]parsedFile, 0, len(files))
	for _, file := range files {
		if err := ctx.Err(); err != nil {
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		finishStage = analysisstats.Measure(ctx, "file_read")
		source, err := os.ReadFile(file)
		finishStage(len(source), err)
		if err != nil {
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		finishStage = analysisstats.Measure(ctx, "parse")
		parsed, err := vbaast.ParseDocument(file, source)
		finishStage(1, err)
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
		finishStage = analysisstats.Measure(ctx, "procedure_ir")
		ir, err := procedureir.BuildParsed(procedureir.BuildOptions{
			RootDir:    a.RootDir,
			Path:       file,
			ModuleKind: moduleKind,
		}, parsed)
		finishStage(len(ir.Procedures), err)
		if err != nil {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, err
		}
		declarationRecovery := false
		if ir.Parse.HasError || ir.Parse.HasMissing {
			if readErr := parsed.Read(func(view vbaast.ParsedView) error {
				declarationRecovery = vbaast.IsDeclarationKeywordRecovery(view.Root, view.Source) ||
					vbaast.IsIdentifierTypeCharacterRecovery(view.Root, view.Source) ||
					vbaast.IsNumericLiteralRecovery(view.Root, view.Source) ||
					lint.IsAcceptedProcedureBoundaryRecovery(view.Root, view.Source, moduleKind)
				return nil
			}); readErr != nil {
				parsed.Close()
				closeParsedFiles(parsedFiles)
				return Result{}, readErr
			}
		}
		if (ir.Parse.HasError || ir.Parse.HasMissing) && !declarationRecovery {
			parsed.Close()
			closeParsedFiles(parsedFiles)
			return Result{}, &ParseError{Path: file, HasError: ir.Parse.HasError, HasMissing: ir.Parse.HasMissing}
		}
		finishStage = analysisstats.Measure(ctx, "cfg")
		controlFlow := vbacfg.BuildDocument(ir)
		finishStage(len(controlFlow.Graphs), nil)
		var intelDocument intel.Document
		if needsTypedExcelAnalysis {
			intelDocument = intel.Document{Path: file, Source: string(source), ModuleKind: moduleKind}
			intelDocument = batchIntelDocument(intelDocument, parsed, ir, controlFlow)
		}
		lines := normalizedSourceLines(string(source))
		var rangeValueConstants map[string]int
		if a.Config.Analyze.DetectRangeValueArrayShape {
			rangeValueConstants = rangeValueModuleIntegerConstants(lines, ir)
		}
		var constantValues map[string]constexpr.Value
		if a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || a.Config.Analyze.DetectObjectArrayComparison || a.Config.Analyze.DetectDeterministicRuntimeErrors {
			constantValues = lint.ConstantValuesFromSource(string(source), &ir, nil)
		}
		parsedFile := parsedFile{
			Path:                      file,
			Lines:                     lines,
			Module:                    strings.TrimSuffix(filepath.Base(file), filepath.Ext(file)),
			ModuleKind:                moduleKind,
			Source:                    source,
			IR:                        ir,
			CFG:                       controlFlow,
			Parsed:                    parsed,
			IntelDocument:             intelDocument,
			RangeValueModuleConstants: rangeValueConstants,
			ConstantValues:            constantValues,
		}
		if a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || a.Config.Analyze.DetectObjectArrayComparison || a.Config.Analyze.DetectDeterministicRuntimeErrors {
			parsedFile.ArrayOptionBase = optionBase(lines)
			parsedFile.ArrayOptionBaseSet = true
			parsedFile.ArrayIntegerModuleConstants = arrayIntegerModuleConstants(parsedFile)
		}
		parsedFiles = append(parsedFiles, parsedFile)
	}
	defer closeParsedFiles(parsedFiles)
	var queryRevision *semanticquery.Revision
	if queryContext.Store != nil {
		revisionID := queryContext.Revision
		if revisionID == "" {
			revisionID = semanticQueryRevisionID(a.RootDir, a.Config, parsedFiles)
		}
		queryRevision = queryContext.Store.Begin(revisionID)
		defer queryRevision.Close()
	}
	recordBatchWorkload(ctx, parsedFiles)
	initializeProjectCapabilityTelemetry(ctx)
	analysis := a
	var warnings []map[string]any
	if needsTypeDB {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityTypeDB)
		finishStage = analysisstats.Measure(ctx, "typedb_load")
		loaded, loadErr := typedb.LoadForRuntime("")
		finishCapability(loadErr)
		finishStage(1, loadErr)
		if loadErr != nil {
			return Result{}, loadErr
		}
		analysis.typeDB = loaded.DB
		analysis.typeDBResolutionIncomplete = !loaded.Complete
		for _, warning := range loaded.Warnings {
			warnings = append(warnings, map[string]any{
				"code": "type_db_load_warning", "message": warning,
			})
		}
	}

	// Resolution is a baseline capability for compile-equivalent diagnostics.
	// Build it before procedure facts so feature classification observes the
	// actual call statuses instead of treating every call as unknown.
	resolutionComplete := analysis.PathFilter == nil && !analysis.typeDBResolutionIncomplete
	for _, file := range parsedFiles {
		if file.IR.Parse.HasError || file.IR.Parse.HasMissing {
			resolutionComplete = false
			break
		}
	}
	finishStage = analysisstats.Measure(ctx, "project_resolution")
	finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityResolution)
	resolutionResolver := buildProjectResolution(parsedFiles, resolutionComplete, analysis.typeDB)
	finishCapability(nil)
	finishStage(1, nil)

	var projectEffects effects.ProjectSummary
	for i := range parsedFiles {
		procedures := sourceProceduresFromIRRef(&parsedFiles[i].IR, parsedFiles[i].CFG)
		parsedFiles[i].Procedures = procedures
		parsedFiles[i].ensureModuleAnalysisFacts()
		if arrayAnalysisEnabled(analysis.Config.Analyze) {
			parsedFiles[i].ArrayVariableCatalog = buildArrayVariableCatalog(parsedFiles[i], parsedFiles[i].moduleDecls())
		}
		parsedFiles[i].moduleFactsFingerprint = semanticModuleFactsFingerprint(parsedFiles[i])
		materializeProcedureAnalysisPlans(&parsedFiles[i], projectEffects, analysis.Config.Analyze)
		recordFactBuilds(ctx, len(procedures))
	}
	capabilityPlan := buildProjectCapabilityPlan(analysis.Config.Analyze, parsedFiles)
	if capabilityPlan.requires(projectCapabilityDataFlowInputs) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityDataFlowInputs)
		for i := range parsedFiles {
			if projectPlansDomain(analysis.Config.Analyze, []parsedFile{parsedFiles[i]}, projectEffects, procedureDomainDataflow) {
				parsedFiles[i].DataFlowModuleBindings = dataFlowBindings(parsedFiles[i].IR.Declarations)
			}
		}
		finishCapability(nil)
	}
	finishStage = analysisstats.Measure(ctx, "effect_summaries")
	if capabilityPlan.requires(projectCapabilityEffects) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityEffects)
		projectEffects = buildProjectEffectsResolved(parsedFiles)
		finishCapability(nil)
		finishStage(projectEffects.ProcedureCount(), nil)
		if recorder := analysisstats.FromContext(ctx); recorder != nil {
			stats := projectEffects.Stats()
			recorder.AddSum("effect_summary_worklist_evaluations", stats.WorklistEvaluations)
			recorder.AddMax("effect_summary_max_propagated_facts_per_procedure", stats.MaxPropagatedFactsPerProcedure)
			recorder.AddSum("effect_summary_total_propagated_facts", stats.TotalPropagatedFacts)
		}
		for i := range parsedFiles {
			materializeProcedureAnalysisPlans(&parsedFiles[i], projectEffects, analysis.Config.Analyze)
		}
		// Keep the revision-owned procedure projection authoritative for the
		// semantic query preparation pass. The worker path still derives a
		// defensive copy, but the prepared effect leaf must come from the same
		// resolved project summary that those workers consume.
		for i := range parsedFiles {
			parsedFiles[i].Procedures = sourceProceduresWithEffects(parsedFiles[i], projectEffects)
		}
		// Effects can add conservative feature seeds (for example, an
		// application-state mutation discovered through a propagated call or a
		// With Application member). Recompute the capability plan before any
		// downstream project indexes are gated so those consumers are retained.
		capabilityPlan = buildProjectCapabilityPlan(analysis.Config.Analyze, parsedFiles)
	} else {
		finishStage(0, nil)
	}
	if err := ctx.Err(); err != nil {
		return Result{}, err
	}
	var cycleFindings []Finding
	if analysisEnabled, known := config.AnalyzeRuleEnabled(a.Config.Analyze, "VBA244"); known && analysisEnabled {
		finishStage = analysisstats.Measure(ctx, "project_wide_diagnostics")
		cycleFindings, err = a.procedureCallCycleFindings(ctx, parsedFiles, projectEffects)
		finishStage(len(cycleFindings), err)
		if err != nil {
			return Result{}, err
		}
	}
	finishStage = analysisstats.Measure(ctx, "object_procedure_summaries")
	var objectAnalysis *objectAnalysisContext
	var finishObjectCapability func(error)
	if a.Config.Analyze.DetectObjectUseBeforeSet && projectPlansDomain(a.Config.Analyze, parsedFiles, projectEffects, procedureDomainObject) {
		finishObjectCapability = beginProjectCapabilityBuild(ctx, projectCapabilityObjectFlow)
		objectAnalysis = buildObjectAnalysisPlans(parsedFiles)
		objectAnalysis.buildSummaries()
		finishStage(len(objectAnalysis.summaries), nil)
	} else {
		finishStage(0, nil)
		if recorder := analysisstats.FromContext(ctx); recorder != nil {
			recorder.AddSum("object_summary_evaluations", 0)
			recorder.AddSum("object_entry_flow_evaluations", 0)
		}
	}
	finishStage = analysisstats.Measure(ctx, "object_entry_states")
	if objectAnalysis != nil {
		objectAnalysis.buildEntryStates()
		finishObjectCapability(nil)
		finishStage(len(objectAnalysis.entries), nil)
		if recorder := analysisstats.FromContext(ctx); recorder != nil {
			recorder.AddSum("object_summary_evaluations", uint64(objectAnalysis.summaryEvaluations))
			recorder.AddSum("object_entry_flow_evaluations", uint64(objectAnalysis.entryFlowEvaluations))
		}
	} else {
		finishStage(0, nil)
	}
	// ProjectConstants is an explicit dependency of the interprocedural array
	// capability. Build it before the project context so the measured capability
	// order follows the same dependency order as the planner closure.
	finishProjectConstants := beginProjectCapabilityBuild(ctx, projectCapabilityProjectConstants)
	analysis.visibleConstants = projectVisibleConstants(parsedFiles, analysis.typeDB)
	analysis.visibleConstantValues = projectConstantValues(parsedFiles, analysis.typeDB)
	finishProjectConstants(nil)
	finishStage = analysisstats.Measure(ctx, "project_context")
	finishArrayCapability := func(error) {}
	if capabilityPlan.requires(projectCapabilityArrayInterprocedural) {
		finishArrayCapability = beginProjectCapabilityBuild(ctx, projectCapabilityArrayInterprocedural)
	}
	analysisCtx := analysis.buildContextWithObjectAnalysisPlan(parsedFiles, objectAnalysis, capabilityPlan, procedureir.ProcedureOnlyResolver(resolutionResolver))
	analysisCtx.queryRevision = queryRevision
	recordArrayInterproceduralTelemetry(ctx, analysisCtx)
	prepareSemanticQueryFacts(analysis, parsedFiles, &analysisCtx)
	finishArrayCapability(nil)
	finishStage(len(analysisCtx.procedures), nil)
	finishStage = analysisstats.Measure(ctx, "project_context_indexes")
	if capabilityPlan.requires(projectCapabilityDictionaryCollection) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityDictionaryCollection)
		analysis.dictionaryCollection = buildDictionaryCollectionIndex(parsedFiles)
		finishCapability(nil)
	}
	if capabilityPlan.requires(projectCapabilityApplicationState) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityApplicationState)
		analysis.applicationStateLeaks = buildApplicationStateLeakIndex(parsedFiles, projectEffects)
		finishCapability(nil)
	}
	if capabilityPlan.requires(projectCapabilityEventReentry) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityEventReentry)
		analysis.eventSafeProcedures = eventSafeProcedures(parsedFiles, projectEffects)
		finishCapability(nil)
	}
	finishStage(1, nil)
	if capabilityPlan.requires(projectCapabilityExcelLoopSymbols) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityExcelLoopSymbols)
		finishStage = analysisstats.Measure(ctx, "project_symbols")
		analysis.excelLoopAccess = buildExcelLoopAccessIndex(parsedFiles, analysis.typeDB, a.RootDir, a.Config)
		finishCapability(nil)
		finishStage(1, nil)
	}
	if needsByRefAnalysis {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityByRefSymbols)
		finishStage = analysisstats.Measure(ctx, "project_symbols")
		byRefSymbolIndex, symbolCount, err := projectByRefSymbolIndex(ctx, a.RootDir, a.Config, a.PathFilter, parsedFiles)
		finishCapability(err)
		finishStage(symbolCount, err)
		if err != nil {
			return Result{}, err
		}
		analysis.byRefSymbolIndex = byRefSymbolIndex
		if recorder := analysisstats.FromContext(ctx); recorder != nil {
			recorder.AddSum("project_symbol_count", uint64(symbolCount))
		}
	}
	findings := cycleFindings
	var resolutionPreflight []Finding
	finishStage = analysisstats.Measure(ctx, "project_wide_diagnostics")
	for i := range parsedFiles {
		resolvedForDiagnostics := procedureir.ResolveView(parsedFiles[i].IR, resolutionResolver)
		for _, diagnostic := range procedureir.DiagnosticsView(resolvedForDiagnostics, resolutionComplete) {
			finding := analysis.resolutionFinding(parsedFiles[i], diagnostic)
			findings = append(findings, finding)
			resolutionPreflight = append(resolutionPreflight, finding)
		}
	}
	finishStage(len(resolutionPreflight), nil)
	var preflightFindings []Finding
	preflightFindings = append(preflightFindings, resolutionPreflight...)
	if capabilityPlan.requires(projectCapabilityExcelAPIHelpers) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityExcelAPIHelpers)
		analysis.errorGuardAliases = projectIsErrorGuardAliases(parsedFiles)
		analysis.errorValueWrappers = projectErrorValueWrappers(parsedFiles)
		finishCapability(nil)
	}
	var publicAPITypeIndex *apiTypeIndex
	if capabilityPlan.requires(projectCapabilityPublicAPITypeIndex) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityPublicAPITypeIndex)
		finishStage = analysisstats.Measure(ctx, "project_symbols")
		publicAPITypeIndex = buildAPITypeIndex(parsedFiles, analysis.typeDB, resolutionComplete)
		finishCapability(nil)
		finishStage(1, nil)
	}
	var analysisMetrics any
	if capabilityPlan.requires(projectCapabilityModuleState) {
		finishCapability := beginProjectCapabilityBuild(ctx, projectCapabilityModuleState)
		finishStage = analysisstats.Measure(ctx, "project_wide_diagnostics")
		moduleState := buildModuleStateAnalysis(a.RootDir, a.Config, parsedFiles)
		finishCapability(nil)
		finishStage(len(moduleState.Findings), nil)
		findings = append(findings, moduleState.Findings...)
		analysisMetrics = moduleState.Metrics
	}
	fileResults, err := analysis.analyzeFilesBounded(ctx, parsedFiles, analysisCtx, projectEffects, publicAPITypeIndex)
	if err != nil {
		return Result{}, err
	}
	for _, fileResult := range fileResults {
		findings = append(findings, fileResult.findings...)
		preflightFindings = append(preflightFindings, fileResult.preflight...)
	}
	finishStage = analysisstats.Measure(ctx, "suppression_and_finalize")
	sortFindings(findings)
	directives, directiveWarnings, err := suppression.DirectivesForFiles(a.RootDir, files)
	if err != nil {
		finishStage(0, err)
		return Result{}, err
	}
	warnings = append(warnings, directiveWarnings...)
	findings, suppressionWarnings := applyInlineSuppressions(findings, directives)
	warnings = append(warnings, suppressionWarnings...)
	finishStage(len(findings), nil)
	result = Result{Findings: findings, Warnings: warnings, AnalysisMetrics: analysisMetrics, PreflightFindings: preflightFindings, AnalyzedFiles: len(parsedFiles)}
	return result, nil
}

func (a Analyzer) analyzeFilesBounded(ctx context.Context, files []parsedFile, analysisCtx analysisContext, projectEffects effects.ProjectSummary, publicAPITypeIndex *apiTypeIndex) ([]parsedFileAnalysisResult, error) {
	return a.analyzeFilesBoundedWith(ctx, files, analysisCtx, projectEffects, publicAPITypeIndex, a.analyzeParsedFileBounded)
}

func (a Analyzer) analyzeFilesBoundedWith(ctx context.Context, files []parsedFile, analysisCtx analysisContext, projectEffects effects.ProjectSummary, publicAPITypeIndex *apiTypeIndex, analyzeFile boundedFileAnalyzer) ([]parsedFileAnalysisResult, error) {
	if len(files) == 0 {
		return nil, nil
	}
	executionLimit := a.analysisWorkerLimit
	if executionLimit <= 0 {
		executionLimit = runtime.GOMAXPROCS(0)
	}
	if executionLimit < 1 {
		executionLimit = 1
	}
	workerLimit := executionLimit
	if workerLimit > len(files) {
		workerLimit = len(files)
	}

	workCtx, cancel := context.WithCancel(ctx)
	defer cancel()
	budget := newAnalysisExecutionBudget(executionLimit)
	for _, file := range files {
		if len(file.Procedures) >= procedureParallelThreshold || len(file.IR.Procedures) >= procedureParallelThreshold {
			if executionLimit > 1 {
				budget.startProcedurePool()
			}
			break
		}
	}
	defer budget.stopProcedurePool()
	workCtx = withAnalysisExecutionBudget(workCtx, budget)
	jobs := make(chan int)
	results := make([]parsedFileAnalysisResult, len(files))
	var (
		workers       sync.WaitGroup
		workerErr     error
		workerErrOnce sync.Once
	)
	workers.Add(workerLimit)
	for worker := 0; worker < workerLimit; worker++ {
		go func() {
			defer workers.Done()
			for {
				select {
				case <-workCtx.Done():
					return
				case index, ok := <-jobs:
					if !ok {
						return
					}
					if err := budget.acquire(workCtx); err != nil {
						return
					}
					fileFindings, filePreflight, err := analyzeFile(workCtx, files[index], analysisCtx, projectEffects, publicAPITypeIndex)
					budget.release()
					results[index] = parsedFileAnalysisResult{findings: fileFindings, preflight: filePreflight, err: err}
					if err != nil {
						workerErrOnce.Do(func() {
							workerErr = err
							cancel()
						})
						return
					}
				}
			}
		}()
	}
sendJobs:
	for index := range files {
		select {
		case <-workCtx.Done():
			break sendJobs
		case jobs <- index:
		}
		if workCtx.Err() != nil {
			break
		}
	}
	close(jobs)
	workers.Wait()
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if workerErr != nil {
		return nil, workerErr
	}
	return results, nil
}

func (a Analyzer) analyzeParsedFileBounded(ctx context.Context, file parsedFile, analysisCtx analysisContext, projectEffects effects.ProjectSummary, publicAPITypeIndex *apiTypeIndex) ([]Finding, []Finding, error) {
	finishStage := analysisstats.Measure(ctx, "procedure_local_diagnostics")
	var findings []Finding
	// Procedure-local analysis only consumes immutable source/IR/CFG facts. Run
	// it outside ParsedDocument.Read so workers never retain or receive a
	// tree-sitter node. The tree is leased below only for the one file-global
	// scan that still needs it.
	file.Root = nil
	procedureFindings, procedureErr := a.analyzeParsedFileContext(ctx, analysisCtx, file, projectEffects, true)
	if procedureErr != nil {
		finishStage(0, procedureErr)
		return nil, nil, procedureErr
	}
	findings = append(findings, procedureFindings...)
	readErr := file.Parsed.Read(func(view vbaast.ParsedView) error {
		if err := ctx.Err(); err != nil {
			return err
		}
		file.Root = view.Root
		if a.Config.Analyze.DetectNonShortCircuitObjectGuard {
			procedures := sourceProceduresWithEffects(file, projectEffects)
			guardFindings, err := a.vba212ScanWithContext(ctx, file, procedures, nil, vba212Context{projectEffects: projectEffects})
			if err != nil {
				return err
			}
			findings = append(findings, guardFindings...)
		}
		if publicAPITypeIndex != nil {
			findings = append(findings, a.publicAPITypeFindings(file, publicAPITypeIndex)...)
		}
		findings = append(findings, a.errorValueWrapperFindings(file)...)
		return nil
	})
	// ParsedView.Root is borrowed for the Read callback only. Do not leave the
	// tree-sitter node reachable from the file projection after the lease ends.
	file.Root = nil
	finishStage(len(findings), readErr)
	if readErr != nil {
		return nil, nil, readErr
	}

	// The typed VBA215/VBA218 analysis uses the same snapshot-owned parsed
	// document. Run it after the tree callback releases its exclusive read
	// lease so the snapshot can reuse that parse without re-entering it.
	finishStage = analysisstats.Measure(ctx, "typed_excel_diagnostics")
	statefulFindings, err := a.statefulExcelCallArgumentFindingsContext(ctx, file)
	if err != nil {
		finishStage(0, err)
		return nil, nil, err
	}
	contractFindings, err := a.excelAPIFailureContractFindingsContext(ctx, file)
	if err != nil {
		finishStage(0, err)
		return nil, nil, err
	}
	findings = append(findings, statefulFindings...)
	findings = append(findings, contractFindings...)
	finishStage(len(statefulFindings)+len(contractFindings), nil)

	finishStage = analysisstats.Measure(ctx, "byref_diagnostics")
	byRefDiagnostics := a.byRefArgumentDiagnosticsContext(ctx, file)
	byRefFindings := a.byRefArgumentFindings(file, byRefDiagnostics.diagnostics)
	findings = append(findings, byRefFindings...)
	finishStage(len(byRefFindings), nil)

	finishStage = analysisstats.Measure(ctx, "compile_equivalent_diagnostics")
	compileFindings, filePreflightFindings := a.compileEquivalentFindingsContext(ctx, file, byRefDiagnostics)
	findings = append(findings, compileFindings...)
	finishStage(len(compileFindings)+len(filePreflightFindings), ctx.Err())
	if err := ctx.Err(); err != nil {
		return nil, nil, err
	}
	return findings, filePreflightFindings, nil
}

func (a Analyzer) byRefArgumentDiagnosticsContext(ctx context.Context, file parsedFile) batchByRefDiagnostics {
	if a.typeDB == nil || ctx.Err() != nil {
		return batchByRefDiagnostics{computed: true}
	}
	diagnostics := (intel.Analyzer{
		RootDir:                    a.RootDir,
		Config:                     a.Config,
		DB:                         a.typeDB,
		TypeDBResolutionIncomplete: a.typeDBResolutionIncomplete,
		WorkspaceSymbolQueryFunc:   a.byRefWorkspaceSymbolQuery,
	}).ByRefArgumentDiagnosticsContext(ctx, file.intelDocument())
	return batchByRefDiagnostics{computed: true, diagnostics: diagnostics}
}

func (a Analyzer) byRefArgumentFindings(file parsedFile, diagnostics []intel.Diagnostic) []Finding {
	if len(diagnostics) == 0 {
		return nil
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != "VBA206" {
			continue
		}
		line := diagnostic.Range.Start.Line + 1
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			candidate := procedures.valueAt(procedureIndex)
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
			"ByRef parameters require a compatible writable argument. Parenthesized and computed arguments can be passed through temporary values, so mutations may not reach the caller.",
			"Pass a compatible writable variable, remove only the argument-level parentheses when mutation is intended, or change the callee parameter to ByVal.",
		)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		out = append(out, finding)
	}
	return out
}

func (a Analyzer) compileEquivalentFindings(file parsedFile) ([]Finding, []Finding) {
	return a.compileEquivalentFindingsContext(context.Background(), file, batchByRefDiagnostics{})
}

func (a Analyzer) compileEquivalentFindingsContext(ctx context.Context, file parsedFile, byRefDiagnostics batchByRefDiagnostics) ([]Finding, []Finding) {
	intelAnalyzer := intel.Analyzer{
		RootDir:                    a.RootDir,
		Config:                     a.Config,
		DB:                         a.typeDB,
		TypeDBResolutionIncomplete: a.typeDBResolutionIncomplete,
		WorkspaceSymbolQueryFunc:   a.byRefWorkspaceSymbolQuery,
		VisibleConstants:           a.visibleConstants,
		ConstantValues:             a.visibleConstantValues,
	}
	var diagnostics []intel.Diagnostic
	if byRefDiagnostics.computed {
		diagnostics = intelAnalyzer.CompileEquivalentDiagnosticsContextWithByRefDiagnostics(ctx, file.intelDocument(), byRefDiagnostics.diagnostics)
	} else {
		diagnostics = intelAnalyzer.CompileEquivalentDiagnosticsContext(ctx, file.intelDocument())
	}
	if len(diagnostics) == 0 {
		return nil, nil
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	preflight := make([]Finding, 0, len(diagnostics))
	for _, diagnostic := range diagnostics {
		metadata, ok := staticrules.Lookup(diagnostic.Code)
		if !ok {
			continue
		}
		line := diagnostic.Range.Start.Line + 1
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			candidate := procedures.valueAt(procedureIndex)
			if line >= candidate.StartLine && line <= candidate.EndLine {
				proc = candidate
				break
			}
		}
		reason, suggestion := compileEquivalentFindingGuidance(diagnostic.Code)
		severity := string(metadata.DefaultSeverity)
		finding := a.simpleFinding(file, proc, line, diagnostic.Code, severity, diagnostic.Message, reason, suggestion)
		finding.Column = diagnostic.Range.Start.Character + 1
		finding.EndLine = diagnostic.Range.End.Line + 1
		finding.EndColumn = diagnostic.Range.End.Character + 1
		if hasRuleSurface(metadata, staticrules.SurfaceAnalyze) {
			out = append(out, finding)
		}
		if metadata.Family == staticrules.FamilyLint && metadata.CompileEquivalent && metadata.PreflightBlocking {
			preflight = append(preflight, finding)
		}
	}
	return out, preflight
}

func hasRuleSurface(metadata staticrules.RuleMetadata, wanted staticrules.RuleSurface) bool {
	for _, surface := range metadata.Surfaces {
		if surface == wanted {
			return true
		}
	}
	return false
}

func compileEquivalentFindingGuidance(code string) (string, string) {
	switch code {
	case "VB060":
		return "VBE rejects assignments to Const declarations.", "Remove the assignment or declare a writable variable instead."
	case "VB061":
		return "VBE rejects a fixed array declaration whose lower bound exceeds its upper bound.", "Use constant bounds with lower less than or equal to upper."
	case "VB037":
		return "VBE rejects Set when the assignment target has a scalar value type.", "Remove Set from the scalar assignment."
	case "VB045":
		return "VBE rejects deterministic argument-count and named-argument binding errors at compile time.", "Pass the required arguments with valid names and do not repeat or reorder named arguments."
	case "VBA228":
		return "VBE rejects a statically incompatible argument passed to a ByRef parameter.", "Pass a variable with the declared parameter type or change the parameter passing contract."
	case "VBA229":
		return "VBE rejects a procedure-local declaration when its As type name cannot be resolved.", "Use a built-in, project-defined, host-library, or referenced TypeLib type name that is available to the project."
	default:
		return "The VBE rejects this deterministic compile-time contract.", "Correct the source so the call or assignment satisfies the VBA compile-time contract."
	}
}

func (a Analyzer) resolutionFinding(file parsedFile, diagnostic procedureir.ResolutionDiagnostic) Finding {
	line := diagnostic.Range.StartLine
	proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
	procedures := file.procedureView()
	for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
		candidate := procedures.valueAt(procedureIndex)
		if line >= candidate.StartLine && line <= candidate.EndLine {
			proc = candidate
			break
		}
	}
	reason, suggestion := compileEquivalentFindingGuidance(diagnostic.Code)
	switch diagnostic.Code {
	case "VB052":
		reason = "The VBE requires a project-local call target to resolve to a callable declaration."
	case "VB053":
		reason = "The VBE cannot choose between multiple visible Enum members without a lexical winner."
	case "VB054":
		reason = "The VBE requires RaiseEvent to name an Event declared in the same object module."
	}
	if shared := procedureir.ResolutionSuggestion(diagnostic.Code); shared != "" {
		suggestion = shared
	}
	finding := a.simpleFinding(file, proc, line, diagnostic.Code, "error", diagnostic.Message, reason, suggestion)
	// procedureir ranges use the same one-based coordinates as vbaast.NodeRange;
	// do not apply the LSP zero-based conversion used by intel diagnostics.
	finding.Column = diagnostic.Range.StartColumn
	finding.EndLine = diagnostic.Range.EndLine
	finding.EndColumn = diagnostic.Range.EndColumn
	return finding
}

// projectByRefSymbolIndex builds the batch project's immutable symbol index
// once. A normal full-project analysis can reuse the already parsed documents;
// a path-filtered analysis intentionally falls back to the complete workspace
// collection because excluded modules remain valid project-local call targets.
func projectByRefSymbolIndex(ctx context.Context, rootDir string, cfg config.Config, pathFilter func(string) bool, parsedFiles []parsedFile) (*intel.WorkspaceResolutionView, int, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, 0, err
	}

	var projectSymbols []intel.Symbol
	if pathFilter != nil {
		var err error
		projectSymbols, err = (intel.Analyzer{RootDir: rootDir, Config: cfg}).WorkspaceSymbolsContext(ctx, nil, "")
		if err != nil {
			return nil, 0, err
		}
	} else {
		symbolAnalyzer := intel.Analyzer{RootDir: rootDir, Config: cfg}
		seen := make(map[string]struct{}, len(parsedFiles))
		for _, file := range parsedFiles {
			if err := ctx.Err(); err != nil {
				return nil, len(projectSymbols), err
			}
			_, included, err := symbols.SourceFileForPath(rootDir, cfg, file.Path)
			if err != nil {
				return nil, len(projectSymbols), err
			}
			if !included {
				continue
			}
			key, err := projectByRefSourcePathKey(file.Path)
			if err != nil {
				return nil, len(projectSymbols), err
			}
			if _, ok := seen[key]; ok {
				continue
			}
			seen[key] = struct{}{}
			fileSymbols, err := symbolAnalyzer.DocumentSymbolsContext(ctx, file.intelDocument())
			if err != nil {
				return nil, len(projectSymbols), err
			}
			for _, symbol := range fileSymbols {
				// WorkspaceSymbols uses IncludeLabels=false. Keep labels out of
				// the batch index so the parsed-document reuse path preserves
				// that historical candidate set.
				switch strings.ToLower(strings.TrimSpace(symbol.Kind)) {
				case "label", "line_number_label":
					continue
				}
				// DocumentSymbolsContext also projects UserForm designer
				// controls as field symbols. The workspace collection used by
				// the old batch path contains source fields, but not those
				// designer-only projections.
				if strings.EqualFold(strings.TrimSpace(symbol.Kind), "field") && symbol.ModuleKind == "" && symbol.Visibility == "" && symbol.Parent == "" {
					continue
				}
				projectSymbols = append(projectSymbols, symbol)
			}
		}
	}
	if err := ctx.Err(); err != nil {
		return nil, len(projectSymbols), err
	}
	return intel.NewWorkspaceResolutionView(projectSymbols), len(projectSymbols), nil
}

func projectByRefSourcePathKey(path string) (string, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	key := filepath.Clean(abs)
	if os.PathSeparator == '\\' {
		key = strings.ToLower(key)
	}
	return key, nil
}

func (a Analyzer) byRefWorkspaceSymbolQuery(_ []intel.Document, query intel.WorkspaceSymbolQuery) ([]intel.Symbol, error) {
	if a.byRefSymbolIndex == nil || (query.Mode != intel.WorkspaceSymbolQueryExact && query.Mode != intel.WorkspaceSymbolQueryQualified) {
		return nil, nil
	}
	return a.byRefSymbolIndex.Query(query), nil
}

func (a Analyzer) statefulExcelCallArgumentFindings(file parsedFile) []Finding {
	out, _ := a.statefulExcelCallArgumentFindingsContext(context.Background(), file)
	return out
}

func (a Analyzer) statefulExcelCallArgumentFindingsContext(ctx context.Context, file parsedFile) ([]Finding, error) {
	if !a.Config.Analyze.DetectStatefulExcelCallArguments || a.typeDB == nil {
		return nil, nil
	}
	diagnostics, err := (intel.Analyzer{RootDir: a.RootDir, Config: a.Config, DB: a.typeDB}).StatefulExcelCallArgumentDiagnosticsContext(ctx, file.intelDocument())
	if err != nil {
		return nil, err
	}
	if len(diagnostics) == 0 {
		return nil, nil
	}
	procedures := file.procedureView()
	out := make([]Finding, 0, len(diagnostics))
	for i, diagnostic := range diagnostics {
		if i&0x3f == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		line := diagnostic.Range.Start.Line + 1
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			candidate := procedures.valueAt(procedureIndex)
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
	return out, ctx.Err()
}

// buildProjectEffectsResolved consumes the already-resolved project revision.
// Keeping resolution outside this function makes the project capability plan
// able to share one resolver with compile-equivalent and procedure-local
// consumers.
func buildProjectEffectsResolved(files []parsedFile) effects.ProjectSummary {
	documents := make([]effects.Document, 0, len(files))
	for _, file := range files {
		documents = append(documents, effects.Document{IR: file.IR, CFG: file.CFG})
	}
	return effects.Build(documents)
}

// buildSingleFileProjectEffects is the explicit boundary for standalone and
// realtime VBA212 callers that do not have a workspace project snapshot. Rule
// implementations receive the resulting value; they never construct a hidden
// project summary themselves.
func buildSingleFileProjectEffects(file parsedFile) effects.ProjectSummary {
	files := []parsedFile{file}
	buildProjectResolution(files, true, nil)
	return buildProjectEffectsResolved(files)
}

func buildProjectResolution(files []parsedFile, complete bool, typeDB *vbadb.DB) procedureir.Resolver {
	resolver := buildResolutionResolver(files, complete, typeDB)
	procedureResolver := procedureir.ProcedureOnlyResolver(resolver)
	for i := range files {
		files[i].IR = procedureir.Resolve(files[i].IR, procedureResolver)
	}
	return resolver
}

func projectVisibleConstants(files []parsedFile, typeDB *vbadb.DB) map[string]bool {
	counts := make(map[string]int)
	add := func(name string) {
		name = strings.ToLower(cleanIdentifier(name))
		if name != "" {
			counts[name]++
		}
	}
	addQualified := func(qualifier, name string) {
		qualifier = strings.ToLower(cleanIdentifier(qualifier))
		name = strings.ToLower(cleanIdentifier(name))
		if qualifier != "" && name != "" {
			counts[qualifier+"."+name]++
		}
	}
	for _, file := range files {
		standardModule := strings.EqualFold(file.ModuleKind, "standard")
		for _, declaration := range file.IR.Declarations {
			if !declaration.IsConst && !strings.EqualFold(declaration.Kind, "enum_member") {
				continue
			}
			if !strings.EqualFold(declaration.Visibility, "public") && !strings.EqualFold(declaration.Visibility, "friend") {
				continue
			}
			if standardModule {
				add(declaration.Name)
			}
			addQualified(file.IR.ModuleName, declaration.Name)
			addQualified(declaration.Parent, declaration.Name)
		}
	}
	if typeDB != nil {
		constants := typeDB.AllConstantsList()
		countsByName := make(map[string]int)
		for _, constant := range constants {
			name := strings.ToLower(cleanIdentifier(constant.Name))
			if name != "" {
				countsByName[name]++
			}
		}
		for _, constant := range constants {
			name := strings.ToLower(cleanIdentifier(constant.Name))
			if name == "" {
				continue
			}
			if countsByName[name] == 1 {
				counts[name]++
			}
			if group := strings.ToLower(cleanIdentifier(constant.EnumGroup)); group != "" {
				counts[group+"."+name]++
			}
			if library := strings.ToLower(cleanIdentifier(constant.Library)); library != "" {
				counts[library+"."+name]++
			}
		}
	}
	constants := make(map[string]bool)
	for name, count := range counts {
		if count == 1 {
			constants[name] = true
		}
	}
	return constants
}

func closeParsedFiles(files []parsedFile) {
	for _, file := range files {
		if file.IntelDocument.Snapshot != nil {
			file.IntelDocument.Snapshot.Retire()
			continue
		}
		if file.Parsed != nil {
			file.Parsed.Close()
		}
	}
}

// intelDocument returns the batch-owned immutable document revision. Batch
// analysis seeds it with the parser already used to build IR and CFG, so the
// typed VBA215 and VBA218 rules share both that parse and their document index.
func (file parsedFile) intelDocument() intel.Document {
	if file.IntelDocument.Snapshot != nil {
		return file.IntelDocument
	}
	return intel.Document{Path: file.Path, Source: string(file.Source), ModuleKind: file.ModuleKind}
}

// batchIntelDocument transfers the parser and already-built immutable
// artifacts into the Intel snapshot used by batch diagnostics. Keeping this
// boundary in one helper makes it explicit that batch preparation and Intel
// analysis operate on the same document revision.
func batchIntelDocument(doc intel.Document, parsed *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) intel.Document {
	doc.Snapshot = intel.NewAnalysisSnapshotWithArtifacts(doc, parsed, intel.AnalysisArtifacts{
		ProcedureIR: ir,
		ControlFlow: controlFlow,
	})
	return doc
}

func SourceNonShortCircuitObjectGuardFindings(rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	return SourceNonShortCircuitObjectGuardFindingsContext(context.Background(), rootDir, path, cfg, source)
}

// SourceNonShortCircuitObjectGuardFindingsContext is the cancellable variant
// of SourceNonShortCircuitObjectGuardFindings.
func SourceNonShortCircuitObjectGuardFindingsContext(ctx context.Context, rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	if !cfg.Analyze.DetectNonShortCircuitObjectGuard {
		return nil, nil
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return nil, err
	}
	defer doc.Close()
	return SourceNonShortCircuitObjectGuardFindingsParsedContext(ctx, rootDir, cfg, doc)
}

// SourceNonShortCircuitObjectGuardFindingsParsed analyzes a caller-owned
// parsed VBA document without closing it or retaining tree-sitter nodes.
func SourceNonShortCircuitObjectGuardFindingsParsed(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	return SourceNonShortCircuitObjectGuardFindingsParsedContext(context.Background(), rootDir, cfg, doc)
}

// SourceNonShortCircuitObjectGuardFindingsParsedContext is the cancellable
// variant of SourceNonShortCircuitObjectGuardFindingsParsed.
func SourceNonShortCircuitObjectGuardFindingsParsedContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	ir, err := procedureir.BuildParsedContext(ctx, procedureir.BuildOptions{RootDir: rootDir}, doc)
	if err != nil {
		return nil, err
	}
	var findings []Finding
	err = doc.ReadContext(ctx, func(view vbaast.ParsedView) error {
		procedures := sourceProceduresFromIRRef(&ir)
		file := parsedFile{
			Path:       view.Path,
			Lines:      normalizedSourceLines(string(view.Source)),
			Module:     strings.TrimSuffix(filepath.Base(view.Path), filepath.Ext(view.Path)),
			Source:     view.Source,
			Root:       view.Root,
			IR:         ir,
			Procedures: procedures,
		}
		file.ensureModuleAnalysisFacts()
		recordFactBuilds(ctx, len(procedures))
		analyzer := Analyzer{RootDir: rootDir, Config: cfg}
		projectEffects := effects.ProjectSummary{}
		if vba212SourceMayHaveGetter(file) {
			projectEffects = buildSingleFileProjectEffects(file)
		}
		var scanErr error
		findings, scanErr = analyzer.vba212ScanWithContext(ctx, file, procedures, nil, vba212Context{projectEffects: projectEffects})
		if scanErr != nil {
			return scanErr
		}
		sortFindings(findings)
		directives, _ := suppression.DirectivesForSource(rootDir, view.Path, string(view.Source))
		findings, _ = applyInlineSuppressions(findings, directives)
		return nil
	})
	return findings, err
}

func SourceRealtimeFindings(rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	return SourceRealtimeFindingsContext(context.Background(), rootDir, path, cfg, source)
}

// SourceRealtimeFindingsContext is the cancellable variant of
// SourceRealtimeFindings.
func SourceRealtimeFindingsContext(ctx context.Context, rootDir, path string, cfg config.Config, source []byte) ([]Finding, error) {
	if !sourceRealtimeAnalysisEnabled(cfg.Analyze) {
		return nil, nil
	}
	doc, err := vbaast.ParseDocument(path, source)
	if err != nil {
		return nil, err
	}
	defer doc.Close()
	return SourceRealtimeFindingsParsedContext(ctx, rootDir, cfg, doc)
}

// SourceRealtimeFindingsParsed runs real-time source analysis against a
// caller-owned parsed VBA document. It does not close doc or retain nodes.
func SourceRealtimeFindingsParsed(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	return SourceRealtimeFindingsParsedContext(context.Background(), rootDir, cfg, doc)
}

// SourceRealtimeFindingsParsedContext is the cancellable variant of
// SourceRealtimeFindingsParsed.
func SourceRealtimeFindingsParsedContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument) ([]Finding, error) {
	ir, err := procedureir.BuildParsedContext(ctx, procedureir.BuildOptions{RootDir: rootDir}, doc)
	if err != nil {
		return nil, err
	}
	return SourceRealtimeFindingsParsedIRContext(ctx, rootDir, cfg, doc, ir)
}

// SourceRealtimeFindingsParsedIR runs real-time source analysis with a
// caller-supplied procedure IR built from doc. It lets immutable LSP snapshots
// reuse their cached IR without reparsing or rewalking procedure syntax.
func SourceRealtimeFindingsParsedIR(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRContext(context.Background(), rootDir, cfg, doc, ir)
}

// SourceRealtimeFindingsParsedIRContext is the cancellable variant of
// SourceRealtimeFindingsParsedIR.
func SourceRealtimeFindingsParsedIRContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR) ([]Finding, error) {
	controlFlow, err := vbacfg.BuildDocumentContext(ctx, ir)
	if err != nil {
		return nil, err
	}
	return SourceRealtimeFindingsParsedIRCFGContext(ctx, rootDir, cfg, doc, ir, controlFlow)
}

// SourceRealtimeFindingsParsedIRCFG runs real-time source analysis with
// caller-supplied procedure IR and control-flow graphs. Immutable LSP
// snapshots use this entry point to reuse both cached analysis layers.
func SourceRealtimeFindingsParsedIRCFG(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGContext(context.Background(), rootDir, cfg, doc, ir, controlFlow)
}

// SourceRealtimeFindingsParsedIRCFGContext is the cancellable variant of
// SourceRealtimeFindingsParsedIRCFG.
func SourceRealtimeFindingsParsedIRCFGContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGWithTypeDBContext(ctx, rootDir, cfg, doc, ir, controlFlow, nil)
}

// SourceRealtimeFindingsParsedIRCFGWithTypeDB runs real-time source analysis
// with an optional caller-owned type database. LSP callers pass their loaded
// database; standalone callers load the built-in database when a typed Excel
// rule is enabled.
func SourceRealtimeFindingsParsedIRCFGWithTypeDB(rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document, typeDB *vbadb.DB) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGWithTypeDBContext(context.Background(), rootDir, cfg, doc, ir, controlFlow, typeDB)
}

// SourceRealtimeFindingsParsedIRCFGWithTypeDBContext is the cancellable
// variant of SourceRealtimeFindingsParsedIRCFGWithTypeDB.
func SourceRealtimeFindingsParsedIRCFGWithTypeDBContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document, typeDB *vbadb.DB) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectContext(ctx, rootDir, cfg, doc, ir, controlFlow, typeDB, effects.ProjectSummary{})
}

// SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectContext adds a
// caller-owned project effect snapshot to the real-time path. LSP Full
// diagnostics use it for cross-file rules while legacy callers retain the
// single-document entry point above.
func SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document, typeDB *vbadb.DB, projectEffects effects.ProjectSummary) ([]Finding, error) {
	return SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectConstantsContext(ctx, rootDir, cfg, doc, ir, controlFlow, typeDB, projectEffects, nil)
}

// SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectConstantsContext is
// the snapshot-aware form used by LSP and other project callers that already
// own a value-bearing constant environment.
func SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectConstantsContext(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document, typeDB *vbadb.DB, projectEffects effects.ProjectSummary, projectConstants map[string]constexpr.Value) ([]Finding, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	ctx, queryContext := withSemanticQueryContext(ctx)
	if !sourceRealtimeAnalysisEnabled(cfg.Analyze) {
		return nil, nil
	}
	if (cfg.Analyze.DetectRangeFindNothingCheck || cfg.Analyze.DetectStatefulExcelCallArguments || cfg.Analyze.DetectExcelAPIFailureContracts || cfg.Analyze.DetectExcelCellAccessInLoops || cfg.Analyze.DetectLoopInvariantExcelObjectResolution || cfg.Analyze.DetectExpensiveFullRangeOperations || cfg.Analyze.DetectValue2PerformanceOpportunities || cfg.Analyze.DetectUntrustedDataFlow || cfg.Analyze.DetectUnsafeCommandConstruction || cfg.Analyze.DetectUnsafeSQLConstruction || cfg.Analyze.DetectUnsafeFilePath) && typeDB == nil {
		var err error
		typeDB, err = vbadb.LoadBuiltin()
		if err != nil {
			return nil, err
		}
	}
	var findings []Finding
	var queryRevision *semanticquery.Revision
	defer func() {
		if queryRevision != nil {
			queryRevision.Close()
		}
	}()
	err := doc.ReadContext(ctx, func(view vbaast.ParsedView) error {
		if err := ctx.Err(); err != nil {
			return err
		}
		moduleKind, classifyErr := realtimeModuleKind(rootDir, cfg, view.Path)
		if classifyErr != nil {
			return classifyErr
		}
		lines := normalizedSourceLines(string(view.Source))
		var rangeValueConstants map[string]int
		if cfg.Analyze.DetectRangeValueArrayShape {
			rangeValueConstants = rangeValueModuleIntegerConstants(lines, ir)
		}
		var constantValues map[string]constexpr.Value
		if cfg.Analyze.DetectArrayLifecycleSafety || cfg.Analyze.DetectRedimPreserveDimension || cfg.Analyze.DetectObjectArrayComparison || cfg.Analyze.DetectDeterministicRuntimeErrors {
			constantValues = lint.ConstantValuesFromSource(string(view.Source), &ir, projectConstants)
		}
		var dataFlowModuleBindings map[string]bool
		if dataFlowInputsEnabled(cfg.Analyze) {
			dataFlowModuleBindings = dataFlowBindings(ir.Declarations)
		}
		procedures := sourceProceduresFromIRRef(&ir, controlFlow)
		file := parsedFile{
			Path:                      view.Path,
			Lines:                     lines,
			Module:                    strings.TrimSuffix(filepath.Base(view.Path), filepath.Ext(view.Path)),
			ModuleKind:                moduleKind,
			Source:                    view.Source,
			Root:                      view.Root,
			IR:                        ir,
			CFG:                       controlFlow,
			Procedures:                procedures,
			RangeValueModuleConstants: rangeValueConstants,
			ConstantValues:            constantValues,
			DataFlowModuleBindings:    dataFlowModuleBindings,
		}
		file.ensureModuleAnalysisFacts()
		file.moduleFactsFingerprint = semanticModuleFactsFingerprint(file)
		materializeProcedureAnalysisPlans(&file, projectEffects, cfg.Analyze)
		recordFactBuilds(ctx, len(procedures))
		if arrayAnalysisEnabled(cfg.Analyze) {
			file.ArrayOptionBase = optionBase(lines)
			file.ArrayOptionBaseSet = true
			file.ArrayIntegerModuleConstants = arrayIntegerModuleConstants(file)
			file.ArrayVariableCatalog = buildArrayVariableCatalog(file, file.moduleDecls())
		}
		if cfg.Analyze.DetectNonShortCircuitObjectGuard && len(projectEffects.AllDirect()) == 0 && vba212SourceMayHaveGetter(file) {
			projectEffects = buildSingleFileProjectEffects(file)
		}
		var excelRootBindings excelRootBindingIndex
		if cfg.Analyze.DetectExcelCellAccessInLoops {
			excelRootBindings = buildRealtimeExcelRootBindingIndex(rootDir, cfg, file)
		}
		analyzer := Analyzer{
			RootDir: rootDir, Config: cfg, typeDB: typeDB,
			visibleConstantValues: projectConstants, excelRootBindings: excelRootBindings,
		}
		if dictionaryCollectionAnalysisEnabled(cfg.Analyze) && projectPlansDomain(cfg.Analyze, []parsedFile{file}, projectEffects, procedureDomainDictionary) {
			analyzer.dictionaryCollection = buildDictionaryCollectionIndex([]parsedFile{file})
		}
		worksheetCodenames := realtimeWorksheetCodenames(rootDir, cfg.Src.Workbook, view.Path)
		procedures = sourceProceduresWithEffects(file, projectEffects)
		moduleDecls := file.moduleDecls()
		findings = append(findings, analyzer.hardcodedSecretFindings(file, procedures)...)
		if len(procedures) == 0 {
			procedures = []sourceProcedure{{StartLine: 1, EndLine: len(file.Lines), StartByte: 0, EndByte: len(file.Source)}}
		}
		contextFiles := []parsedFile{file}
		if queryContext.Store != nil && queryRevision == nil {
			revisionID := queryContext.Revision
			if revisionID == "" {
				revisionID = semanticQueryRevisionID(rootDir, cfg, contextFiles)
			}
			queryRevision = queryContext.Store.Begin(revisionID)
		}
		analysisCtx := analyzer.buildContext(contextFiles)
		analysisCtx.queryRevision = queryRevision
		// buildContext materializes the participant-restricted plan after the
		// initial procedure projection above was copied. Rebind the realtime
		// projection to that materialized revision so batch and realtime execute
		// the same array participant boundary.
		file = contextFiles[0]
		materializedProcedures := sourceProceduresWithEffects(file, projectEffects)
		procedures = rebindRealtimeProcedureProjection(procedures, materializedProcedures)
		// The prepared query facts are built from the context-file projection;
		// attach the effect-bearing materialized procedures before preparation so
		// callee effect leaves and procedure effect summaries stay coherent with
		// the findings path below.
		contextFiles[0].Procedures = procedures
		recordArrayInterproceduralTelemetry(ctx, analysisCtx)
		analysisCtx.projectEffects = projectEffects
		prepareSemanticQueryFacts(analyzer, contextFiles, &analysisCtx)
		// prepareSemanticQueryFacts attaches the immutable projection to the
		// context-file slice; refresh the value copy used by realtime workers.
		file = contextFiles[0]
		for i, proc := range procedures {
			if i&0x1f == 0 {
				if err := ctx.Err(); err != nil {
					return err
				}
			}
			procedureFindings, err := analyzer.sourceRealtimeProcedureFindingsContext(ctx, file, proc, moduleDecls, worksheetCodenames, analysisCtx)
			if err != nil {
				return err
			}
			findings = append(findings, procedureFindings...)
		}
		if cfg.Analyze.DetectNonShortCircuitObjectGuard {
			guardFindings, err := analyzer.vba212ScanWithContext(ctx, file, procedures, nil, vba212Context{projectEffects: projectEffects})
			if err != nil {
				return err
			}
			findings = append(findings, guardFindings...)
		}
		if err := ctx.Err(); err != nil {
			return err
		}
		statefulFindings, err := analyzer.statefulExcelCallArgumentFindingsContext(ctx, file)
		if err != nil {
			return err
		}
		findings = append(findings, statefulFindings...)
		contractFindings, err := analyzer.excelAPIFailureContractFindingsContext(ctx, file)
		if err != nil {
			return err
		}
		findings = append(findings, contractFindings...)
		wrapperFindings, err := analyzer.errorValueWrapperFindingsContext(ctx, file)
		if err != nil {
			return err
		}
		findings = append(findings, wrapperFindings...)
		findings = realtimeFindings(findings)
		sortFindings(findings)
		directives, _ := suppression.DirectivesForSource(rootDir, view.Path, string(view.Source))
		findings, _ = applyInlineSuppressions(findings, directives)
		return nil
	})
	return findings, err
}

// VBA206 is evaluated by intel.Diagnostics after this callback so the LSP can
// resolve the latest workspace-document overlays through its symbol provider.
var sourceRealtimeRuleIDs = []string{"VBA201", "VBA204", "VBA206", "VBA208", "VBA209", "VBA212", "VBA213", "VBA215", "VBA216", "VBA217", "VBA218", "VBA219", "VBA223", "VBA224", "VBA225", "VBA226", "VBA227", "VBA228", "VBA229", "VBA230", "VBA231", "VBA232", "VBA233", "VBA234", "VBA235", "VBA236", "VBA237", "VBA238", "VBA239", "VBA241", "VBA242", "VBA243", "VBA245", "VBA246", "VBA247", "VBA248", "VBA249"}

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

func realtimeModuleKind(rootDir string, cfg config.Config, path string) (string, error) {
	sourceFile, included, err := symbols.SourceFileForPath(rootDir, cfg, path)
	if err != nil {
		return "", err
	}
	if !included {
		return "", nil
	}
	return sourceFile.ModuleKind, nil
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

func (a Analyzer) sourceRealtimeProcedureFindingsContext(ctx context.Context, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, worksheetCodenames map[string]string, analysisCtx analysisContext) ([]Finding, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	decls := newDeclarationScope(file, proc)
	decls.module = moduleDecls
	findAssignments := map[string]rangeFindAssignmentInfo{}
	guardedFinds := map[string]bool{}
	withStack := make([]withInfo, 0)
	worksheetRoots := newWorksheetRootTracker(worksheetCodenames)
	var findings []Finding
	plan := proc.analysisPlan(a.Config.Analyze, moduleDecls)
	profile := newProcedureDomainProfile(ctx)
	profile.plannerDecisions(plan)
	profile.dataflowLaneDecisions(plan)
	profile.realtimePlanSummary(plan)
	defer profile.flush()
	var resultStore procedureSemanticResultStore
	resultStore.queryRevision = analysisCtx.queryRevision
	var arrayResult *ArrayAnalysisResult
	if plan.runsKernel(procedureKernelArray) {
		var err error
		arrayResult, err = resultStore.materializeArray(ctx, a, file, proc, analysisCtx, moduleDecls, plan)
		if err != nil {
			return nil, err
		}
		if arrayResult != nil && !resultStore.arrayHit {
			profile.kernel()
			profile.add(analysisstats.CounterArrayKernelRuns, 1)
			profile.add(analysisstats.CounterArrayCFGWalks, arrayResult.cfgWalks)
			profile.add(analysisstats.CounterArrayProjectionRuns, arrayResult.projectionRuns)
			if arrayResult.projectionRuns > 0 {
				profile.add(analysisstats.CounterArrayCandidateProcedures, 1)
			}
		}
	}
	findings = append(findings, a.opaqueBooleanArgumentFindings(file, proc, analysisCtx.procedures)...)
	if plan.runsProjection(procedureProjectionRuntime) {
		findings = append(findings, a.deterministicRuntimeErrorFindingsWithArrayResult(file, proc, analysisCtx, moduleDecls, resultStore.arrayProjection(profile))...)
	}
	if plan.runsAnyProjection(
		procedureProjectionDictionaryGuard,
		procedureProjectionDictionaryCompareMode,
		procedureProjectionDictionaryLoop,
		procedureProjectionDictionaryKeyNormalization,
		procedureProjectionDictionaryLateBound,
		procedureProjectionDictionaryCollectionMutation,
		procedureProjectionDictionaryIndexOrigin,
	) && dictionaryCollectionAnalysisEnabled(a.Config.Analyze) {
		dictionaryFindings, err := a.dictionaryCollectionSafetyFindings(ctx, file, proc, moduleDecls)
		if err != nil {
			return nil, err
		}
		findings = append(findings, dictionaryFindings...)
	}
	if plan.runsProjection(procedureProjectionDictionaryIteration) && a.Config.Analyze.DetectDictionaryIterationValueUsage {
		findings = append(findings, a.dictionaryIterationValueUsageFindings(file, proc, moduleDecls)...)
	}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		if i&0xff == 0 {
			if err := ctx.Err(); err != nil {
				return nil, err
			}
		}
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		worksheetStmt, worksheetStatementStart := worksheetLogicalStatement(file.Lines, i, proc.EndLine-1)
		if stmt == "" {
			continue
		}
		if endWithRe.MatchString(stmt) {
			if len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
			worksheetRoots.popWith()
			continue
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
		if setAssignRe.MatchString(stmt) {
			if name, receiver, ok := rangeFindAssignment(stmt, currentWithExpression(withStack)); ok {
				findAssignments[strings.ToLower(name)] = rangeFindAssignmentInfo{Line: lineNo, Receiver: receiver}
			}
			continue
		}
		if a.Config.Analyze.DetectRangeFindNothingCheck {
			findings = append(findings, a.rangeFindFindings(file, proc, lineNo, stmt, findAssignments, guardedFinds)...)
		}
		if plan.runsProjection(procedureProjectionArrayComparison) && objectNothingEqualityLineIsExecutable(proc, lineNo, stmt) {
			arrayMeasurement := profile.begin(procedureDomainArray)
			arrayFindings := a.objectArrayComparisonFindings(file, proc, lineNo, stmt, decls)
			profile.kernel()
			profile.add(analysisstats.CounterArrayCandidateProcedures, 1)
			arrayMeasurement.finish(len(arrayFindings))
			findings = append(findings, arrayFindings...)
		}
	}
	if plan.runsProjection(procedureProjectionArrayRangeShape) {
		if result := resultStore.arrayProjection(profile); result != nil {
			findings = append(findings, result.rangeShape()...)
		} else {
			findings = append(findings, a.rangeValueShapeFindings(file, proc)...)
		}
	}
	if plan.runsAnyProjection(procedureProjectionArrayLifecycle, procedureProjectionArrayRedim, procedureProjectionArrayComparison) {
		if result := resultStore.arrayProjection(profile); result != nil {
			findings = append(findings, result.lifecycle()...)
		} else {
			findings = append(findings, a.arrayLifecycleFindings(file, proc, analysisCtx, moduleDecls)...)
		}
		findings = suppressDeterministicArrayWarningDuplicates(findings)
	}
	if plan.runsProjection(procedureProjectionArrayRedimLoop) {
		if result := resultStore.arrayProjection(profile); result != nil {
			findings = append(findings, result.redim()...)
		} else {
			findings = append(findings, a.redimPreserveLoopFindings(file, proc, moduleDecls)...)
		}
	}
	if plan.runsProjection(procedureProjectionErrorHandler) && a.Config.Analyze.DetectErrorHandlerFallthrough {
		findings = append(findings, a.errorHandlerFallthroughFindings(file, proc)...)
	}
	if plan.runsProjection(procedureProjectionErrorSuppression) && a.Config.Analyze.DetectErrorSuppressionPropagation {
		findings = append(findings, a.errorSuppressionFindings(file, proc, analysisCtx.projectEffects)...)
	}
	var realtimeHTTPFindings []Finding
	if plan.runsDataflowLane(procedureDataflowLaneGeneric) || plan.runsDataflowLane(procedureDataflowLaneHTTP) {
		dataFlowFindings, httpFindings, err := a.executeDataflowLanes(ctx, file, proc, plan, &resultStore, profile, nil)
		if err != nil {
			return nil, err
		}
		findings = append(findings, dataFlowFindings...)
		realtimeHTTPFindings = httpFindings
	}
	if plan.runsProjection(procedureProjectionDataflowFile) {
		findings = append(findings, a.filePathSafetyFindings(file, proc)...)
	}
	findings = append(findings, realtimeHTTPFindings...)
	if plan.runsProjection(procedureProjectionResource) && a.Config.Analyze.DetectResourceLeaks {
		findings = append(findings, a.resourceLeakFindings(file, proc)...)
	}
	if plan.runsProjection(procedureProjectionExcelLoop) && a.Config.Analyze.DetectExcelCellAccessInLoops {
		findings = append(findings, a.excelLoopAccessFindings(file, proc)...)
	}
	if plan.runsProjection(procedureProjectionExcelInvariant) && a.Config.Analyze.DetectLoopInvariantExcelObjectResolution {
		findings = append(findings, a.excelLoopInvariantFindings(file, proc)...)
	}
	if plan.runsProjection(procedureProjectionExcelRange) && a.Config.Analyze.DetectExpensiveFullRangeOperations {
		findings = append(findings, a.expensiveFullRangeOperationFindings(file, proc)...)
	}
	if plan.runsProjection(procedureProjectionExcelValue2) && a.Config.Analyze.DetectValue2PerformanceOpportunities {
		findings = append(findings, a.value2PerformanceFindings(file, proc)...)
	}
	return findings, ctx.Err()
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

func buildResolutionResolver(files []parsedFile, complete bool, typeDB *vbadb.DB) procedureir.Resolver {
	resolverSymbols := make([]procedureir.ResolverSymbol, 0)
	for _, file := range files {
		module := strings.TrimSpace(file.IR.ModuleName)
		if module == "" {
			module = file.Module
		}
		for _, declaration := range file.IR.Declarations {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: declaration.Name, Type: declaration.Type, Module: module, ModuleKind: file.IR.ModuleKind,
				Kind: declaration.Kind, Visibility: declaration.Visibility, File: file.IR.Path,
				Line: declaration.Range.StartLine, Parent: declaration.Parent, Recovered: declaration.Recovered,
				IsArray: declaration.IsArray, IsConst: declaration.IsConst,
				ValueShape:          declaration.ValueShape,
				ConditionalBranches: append([]procedureir.ConditionalBranch(nil), declaration.ConditionalBranches...),
			})
		}
		for _, procedure := range file.IR.Procedures {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: procedure.Symbol.Name, Type: procedure.Symbol.ReturnType, Module: module, ModuleKind: file.IR.ModuleKind,
				Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility, File: file.IR.Path,
				Line: procedure.Symbol.DeclarationRange.StartLine, Recovered: procedure.Symbol.Recovered,
				IsArray: procedure.Symbol.IsArray, ValueShape: procedure.Symbol.ValueShape,
				ConditionalBranches: append([]procedureir.ConditionalBranch(nil), procedure.Symbol.ConditionalBranches...),
			})
		}
	}
	if typeDB != nil {
		for _, constant := range typeDB.AllConstantsList() {
			if strings.TrimSpace(constant.EnumGroup) == "" {
				continue
			}
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: constant.Name, Parent: constant.EnumGroup, Module: constant.Library,
				ModuleKind: "external", Kind: "enum_member", Visibility: "Public",
				File: "<typelib>" + constant.Library, Line: 0,
				IsConst: true, ValueShape: procedureir.ValueShapeScalar,
			})
		}
	}
	return procedureir.NewResolverWithCompleteness(resolverSymbols, complete)
}

func (a Analyzer) buildContext(files []parsedFile) analysisContext {
	return a.buildContextWithObjectAnalysisPlan(files, nil, buildProjectCapabilityPlan(a.Config.Analyze, files), nil)
}

func (a Analyzer) buildContextWithObjectAnalysisPlan(files []parsedFile, objectAnalysis *objectAnalysisContext, capabilityPlan projectCapabilityPlan, projectResolver procedureir.Resolver) analysisContext {
	ctx := analysisContext{
		functionReturns:                  map[string]string{},
		functionShapes:                   map[string]procedureir.ValueShapeKind{},
		functionNamesSeen:                map[string]bool{},
		functionAmbiguous:                map[string]bool{},
		arrayReturns:                     map[string]arrayValue{},
		arrayAllocationGuards:            map[string]bool{},
		arrayByRefAllocations:            arrayByRefAllocationSummaries{},
		arrayByRefConditionalAllocations: arrayByRefConditionalAllocations{},
		arrayByRefLengthAllocations:      arrayByRefLengthAllocations{},
		arrayModuleAllocations:           arrayModuleAllocationSummaries{},
		arrayModuleConfigurations:        map[string]arrayModuleConfigurationState{},
		arrayModuleEntryStates:           arrayModuleEntryStates{},
		arrayPrivateTargets:              map[string]sourceProcedure{},
		arrayStats:                       &arrayInterproceduralStats{strategy: a.arrayStrategy},
		arrayByRefEntryStates:            map[string]map[int]bool{},
		arrayByRefEntryConditions:        map[string]map[int]string{},
		procedures:                       map[string]procedureSignature{},
		objectAnalysis:                   objectAnalysis,
		worksheetCodenames:               map[string]string{},
	}
	var resolverSymbols []procedureir.ResolverSymbol
	if projectResolver == nil {
		resolverSymbols = make([]procedureir.ResolverSymbol, 0)
	}
	workbookRoot := filepath.Clean(filepath.Join(a.RootDir, a.Config.Src.Workbook))
	for _, file := range files {
		if rel, err := filepath.Rel(workbookRoot, file.Path); err == nil && rel != "" && !strings.HasPrefix(rel, "..") && !strings.EqualFold(file.Module, "ThisWorkbook") {
			ctx.worksheetCodenames[strings.ToLower(file.Module)] = file.Module
		}
		module := strings.TrimSpace(file.IR.ModuleName)
		if module == "" {
			module = file.Module
		}
		if projectResolver == nil {
			for _, procedure := range file.IR.Procedures {
				resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
					Name:       procedure.Symbol.Name,
					Module:     module,
					ModuleKind: file.IR.ModuleKind,
					Kind:       string(procedure.Symbol.Kind),
					Visibility: procedure.Symbol.Visibility,
					File:       file.IR.Path,
					Line:       procedure.Symbol.DeclarationRange.StartLine,
				})
			}
		}
		procedures := file.procedureView()
		for procedureIndex := 0; procedureIndex < procedures.Len(); procedureIndex++ {
			proc := procedures.valueAt(procedureIndex)
			functionName := strings.ToLower(proc.Name)
			if functionName != "" {
				if ctx.functionNamesSeen[functionName] {
					// A receiver-less call with duplicate project candidates is
					// ambiguous. Never infer a shape from declaration order.
					delete(ctx.functionShapes, functionName)
					delete(ctx.functionReturns, functionName)
					ctx.functionAmbiguous[functionName] = true
				} else {
					ctx.functionNamesSeen[functionName] = true
					if isObjectType(proc.ReturnType) {
						ctx.functionReturns[functionName] = proc.ReturnType
					}
				}
			}
			// Only built-in scalar return types are strong enough to reject a
			// For Each source. A user-defined class/UDT may expose an enumerator
			// and therefore remains unknown even when its IR shape is scalar.
			if functionName == "" || ctx.functionAmbiguous[functionName] {
				continue
			}
			if proc.ReturnValueShape == procedureir.ValueShapeScalar && arrayKnownScalarType(proc.ReturnType) && !isObjectType(proc.ReturnType) {
				ctx.functionShapes[functionName] = proc.ReturnValueShape
			}
			signature := procedureSignature{
				Name:       proc.Name,
				ReturnType: proc.ReturnType,
				Params:     proc.Params,
			}
			ctx.procedures[strings.ToLower(proc.Name)] = signature
			ctx.procedures[strings.ToLower(file.Module+"."+proc.Name)] = signature
		}
	}
	ctx.procedureResolver = projectResolver
	if ctx.procedureResolver == nil {
		ctx.procedureResolver = procedureir.NewResolver(resolverSymbols)
	}
	// A completely scalar project with no uncertain calls cannot produce an
	// array diagnostic. Avoid building the interprocedural array indexes in
	// that proven-negative case; any call or incomplete fact fails this gate
	// open through buildProcedureAnalysisPlan.
	if capabilityPlan.requires(projectCapabilityArrayInterprocedural) {
		ctx.arrayAllocationGuards = inferArrayAllocationGuards(files)
		ctx.arrayPrivateTargets = arrayPrivateProcedureTargets(files)
		ctx.arrayParticipants, ctx.arrayInterproceduralParticipants, ctx.arrayParticipantKeys = buildArrayParticipantSets(files, ctx)
		materializeArrayParticipantPlans(files, a.Config.Analyze, ctx.arrayParticipants, ctx.arrayParticipantKeys)
		ctx.arrayReturns = inferArrayReturnSummaries(files, ctx.arrayAllocationGuards, ctx)
		ctx.arrayByRefAllocations = inferArrayByRefAllocationSummaries(files, ctx, ctx.arrayPrivateTargets)
		ctx.arrayByRefConditionalAllocations = inferArrayByRefConditionalAllocations(files)
		ctx.arrayByRefLengthAllocations = inferArrayByRefLengthAllocations(files)
		ctx.arrayModuleAllocations = inferArrayModuleAllocationSummaries(files, ctx, ctx.arrayPrivateTargets, ctx.arrayByRefAllocations)
		ctx.arrayModuleConfigurations = inferArrayModuleConfigurationStates(files, ctx.arrayModuleAllocations)
		ctx.arrayModuleEntryStates = inferArrayModuleEntryStates(a, files, ctx)
		ctx.arrayByRefEntryStates, ctx.arrayByRefEntryConditions = inferArrayByRefEntryStates(a, files, ctx)
	}
	return ctx
}

func recordBatchWorkload(ctx context.Context, files []parsedFile) {
	recorder := analysisstats.FromContext(ctx)
	if recorder == nil {
		return
	}
	recorder.AddSum("file_count", uint64(len(files)))
	// Seed maximum dimensions so an empty project still reports a complete
	// workload shape with zero values.
	for _, name := range []string{
		"max_lines_per_file", "max_procedures_per_file", "max_calls_per_file",
		"max_statements_per_procedure", "max_cfg_blocks_per_procedure",
		"max_cfg_edges_per_procedure",
	} {
		recorder.AddMax(name, 0)
	}
	for _, file := range files {
		lineCount := physicalSourceLineCount(file.Lines)
		recorder.AddSum("line_count", uint64(lineCount))
		recorder.AddSum("module_declaration_count", uint64(len(file.IR.Declarations)))
		procedureCount := len(file.IR.Procedures)
		recorder.AddSum("procedure_count", uint64(procedureCount))
		recorder.AddMax("max_lines_per_file", uint64(lineCount))
		recorder.AddMax("max_procedures_per_file", uint64(procedureCount))
		callsPerFile := 0
		for _, procedure := range file.IR.Procedures {
			statementCount := len(procedure.Statements)
			recorder.AddSum("statement_count", uint64(statementCount))
			recorder.AddSum("expression_count", uint64(len(procedure.Expressions)))
			callsPerFile += len(procedure.Calls)
			recorder.AddSum("call_site_count", uint64(len(procedure.Calls)))
			recorder.AddMax("max_statements_per_procedure", uint64(statementCount))
		}
		recorder.AddMax("max_calls_per_file", uint64(callsPerFile))
		for _, graph := range file.CFG.Graphs {
			recorder.AddMax("max_cfg_blocks_per_procedure", uint64(len(graph.Blocks)))
			recorder.AddMax("max_cfg_edges_per_procedure", uint64(len(graph.Edges)))
			recorder.AddSum("cfg_block_count", uint64(len(graph.Blocks)))
			recorder.AddSum("cfg_edge_count", uint64(len(graph.Edges)))
		}
	}
}

func recordArrayInterproceduralTelemetry(ctx context.Context, analysisCtx analysisContext) {
	recorder := analysisstats.FromContext(ctx)
	if recorder == nil || analysisCtx.arrayParticipants == nil {
		return
	}
	recorder.AddSum(analysisstats.ArrayParticipantProceduresCounter, uint64(len(analysisCtx.arrayParticipants)))
	if analysisCtx.arrayStats != nil {
		cfgWalks, revisits, compactWalks, legacyWalks, fallbackWalks, fallbackEmptyState, fallbackIndex, fallbackUnsupported := analysisCtx.arrayStats.snapshot()
		recorder.AddSum(analysisstats.ArrayInterproceduralCFGWalksCounter, cfgWalks)
		recorder.AddSum(analysisstats.ArrayWorklistRevisitsCounter, revisits)
		recorder.AddSum("array_compact_cfg_walks", compactWalks)
		recorder.AddSum("array_legacy_cfg_walks", legacyWalks)
		recorder.AddSum("array_cfg_fallbacks", fallbackWalks)
		recorder.AddSum("array_cfg_fallback_empty_state", fallbackEmptyState)
		recorder.AddSum("array_cfg_fallback_index", fallbackIndex)
		recorder.AddSum("array_cfg_fallback_unsupported", fallbackUnsupported)
	}
}

func (a Analyzer) analyzeParsedFileContext(cancelCtx context.Context, ctx analysisContext, file parsedFile, projectEffects effects.ProjectSummary, filePermitHeld bool) ([]Finding, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	file.ensureModuleAnalysisFacts()
	var findings []Finding
	procedures := sourceProceduresWithEffects(file, projectEffects)
	moduleDecls := file.moduleDecls()
	findings = append(findings, a.hardcodedSecretFindings(file, procedures)...)
	procedureFindings, err := a.analyzeProcedureContextsBounded(cancelCtx, file, procedures, moduleDecls, ctx, projectEffects, filePermitHeld)
	if err != nil {
		return nil, err
	}
	findings = append(findings, procedureFindings...)
	return findings, cancelCtx.Err()
}

// analyzeProcedureContextsBounded evaluates procedure-local rules using the
// shared analysis execution budget. Each result is stored by source index and
// merged in that order, so worker completion order cannot affect findings.
func (a Analyzer) analyzeProcedureContextsBounded(cancelCtx context.Context, file parsedFile, procedures []sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, projectEffects effects.ProjectSummary, filePermitHeld bool) ([]Finding, error) {
	file.ensureModuleAnalysisFacts()
	if moduleDecls == nil {
		moduleDecls = file.moduleDecls()
	}
	if len(procedures) == 0 {
		proc := sourceProcedure{StartLine: 1, EndLine: len(file.Lines)}
		profile := newProcedureDomainProfile(cancelCtx)
		findings, err := a.analyzeProcedureContext(cancelCtx, file, proc, moduleDecls, ctx, projectEffects, map[string]bool{}, profile)
		profile.flush()
		return findings, err
	}
	budget := analysisExecutionBudgetFromContext(cancelCtx)
	if !filePermitHeld || budget == nil || budget.procedureJobs == nil || budget.limit < 2 || len(procedures) < procedureParallelThreshold {
		profile := newProcedureDomainProfile(cancelCtx)
		reportedMissingHelpers := map[string]bool{}
		var findings []Finding
		for _, proc := range procedures {
			if err := cancelCtx.Err(); err != nil {
				profile.flush()
				return nil, err
			}
			procedureFindings, err := a.analyzeProcedureContext(cancelCtx, file, proc, moduleDecls, ctx, projectEffects, reportedMissingHelpers, profile)
			if err != nil {
				profile.flush()
				return nil, err
			}
			findings = append(findings, procedureFindings...)
		}
		profile.flush()
		return findings, nil
	}

	// The enclosing file job owns one permit. Yield it before waiting for
	// procedure jobs; otherwise a full set of file workers could deadlock while
	// trying to acquire the same budget for their children.
	// The caller explicitly confirms that it owns the file-level permit. A
	// future caller without that permit must take the serial path above rather
	// than releasing an unrelated semaphore slot.
	budget.release()
	filePermitHeld = false
	defer func() {
		// The caller owns a permit for the duration of analyzeParsedFileBounded.
		// Cancellation is handled by the caller's existing error path. Reacquire
		// without that cancellation so the ownership invariant remains balanced
		// before the file worker releases its original permit.
		if !filePermitHeld {
			_ = budget.acquire(context.Background())
		}
	}()

	workCtx, cancel := context.WithCancel(cancelCtx)
	defer cancel()
	results := make([][]Finding, len(procedures))
	var (
		batches   sync.WaitGroup
		firstErr  error
		errOnce   sync.Once
		startOnce sync.Once
	)
	batchSize := (len(procedures) + budget.limit*4 - 1) / (budget.limit * 4)
	if batchSize < 1 {
		batchSize = 1
	}
	for start := 0; start < len(procedures); start += batchSize {
		end := start + batchSize
		if end > len(procedures) {
			end = len(procedures)
		}
		batchStart, batchEnd := start, end
		batches.Add(1)
		job := procedureBatchJob{
			ctx: workCtx,
			run: func(jobCtx context.Context) ([]Finding, error) {
				procedureFile := file
				procedureFile.Root = nil
				profile := newProcedureDomainProfile(jobCtx)
				startOnce.Do(func() {
					if a.procedureAnalysisStartHook != nil {
						a.procedureAnalysisStartHook()
					}
				})
				for index := batchStart; index < batchEnd; index++ {
					if err := jobCtx.Err(); err != nil {
						profile.flush()
						return nil, err
					}
					findings, err := a.analyzeProcedureContext(jobCtx, procedureFile, procedures[index], moduleDecls, ctx, projectEffects, map[string]bool{}, profile)
					if err != nil {
						profile.flush()
						return nil, err
					}
					results[index] = findings
				}
				profile.flush()
				return nil, nil
			},
			complete: func(_ []Finding, err error) {
				defer batches.Done()
				if err != nil && !errors.Is(err, context.Canceled) && !errors.Is(err, context.DeadlineExceeded) {
					errOnce.Do(func() {
						firstErr = err
						cancel()
					})
				}
			},
		}
		if !budget.submitProcedureBatch(workCtx, job) {
			batches.Done()
			break
		}
	}
	batches.Wait()
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	if firstErr != nil {
		return nil, firstErr
	}

	// Helper dependency findings are file-global one-time diagnostics. Workers
	// collect them locally, then this source-order merge keeps the first
	// occurrence and drops later duplicates exactly like the serial map did.
	seenHelpers := map[string]bool{}
	var findings []Finding
	for index := range results {
		for _, finding := range results[index] {
			if finding.Code == "VBA105" || finding.Code == "VBA106" {
				if seenHelpers[finding.Code] {
					continue
				}
				seenHelpers[finding.Code] = true
			}
			findings = append(findings, finding)
		}
	}
	return findings, nil
}

func (a Analyzer) analyzeProcedureContext(cancelCtx context.Context, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, projectEffects effects.ProjectSummary, reportedMissingHelpers map[string]bool, profile *procedureDomainProfile) ([]Finding, error) {
	plan := proc.analysisPlan(a.Config.Analyze, moduleDecls)
	return a.executeProcedureAnalysisPlan(cancelCtx, file, proc, moduleDecls, ctx, projectEffects, reportedMissingHelpers, profile, plan)
}

// executeProcedureAnalysisPlan evaluates one immutable procedure plan. The
// body keeps the established canonical finding append order; all semantic
// preparation is selected through plan kernel/projection masks, so callers
// never derive execution from worker completion order or a map iteration.
func (a Analyzer) executeProcedureAnalysisPlan(cancelCtx context.Context, file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration, ctx analysisContext, projectEffects effects.ProjectSummary, reportedMissingHelpers map[string]bool, profile *procedureDomainProfile, plan procedureAnalysisPlan) ([]Finding, error) {
	if err := cancelCtx.Err(); err != nil {
		return nil, err
	}
	decls := newDeclarationScope(file, proc)
	decls.module = moduleDecls
	shadowedVBA205 := vba205ShadowedIdentifiersWithFacts(proc, decls, ctx, file.moduleAnalysisFacts())
	withStack := make([]withInfo, 0)
	findAssignments := map[string]rangeFindAssignmentInfo{}
	guardedFinds := map[string]bool{}
	worksheetRoots := newWorksheetRootTracker(ctx.worksheetCodenames)
	var findings []Finding
	var candidateCounters uint64
	profile.plannerDecisions(plan)
	profile.dataflowLaneDecisions(plan)
	profile.planSummary(plan)
	var resultStore procedureSemanticResultStore
	resultStore.queryRevision = ctx.queryRevision
	var arrayResult *ArrayAnalysisResult
	if plan.runsKernel(procedureKernelArray) {
		var err error
		arrayResult, err = resultStore.materializeArray(cancelCtx, a, file, proc, ctx, moduleDecls, plan)
		if err != nil {
			return nil, err
		}
		if arrayResult != nil && !resultStore.arrayHit {
			profile.kernel()
			profile.add(analysisstats.CounterArrayKernelRuns, 1)
			profile.add(analysisstats.CounterArrayCFGWalks, arrayResult.cfgWalks)
			profile.add(analysisstats.CounterArrayProjectionRuns, arrayResult.projectionRuns)
			if arrayResult.projectionRuns > 0 {
				profile.candidate(&candidateCounters, analysisstats.CounterArrayCandidateProcedures)
			}
		}
	}
	otherMeasurement := profile.begin(procedureDomainOther)
	opaqueFindings := a.opaqueBooleanArgumentFindings(file, proc, ctx.procedures)
	if a.Config.Analyze.DetectOpaqueBooleanArguments {
		profile.kernel()
	}
	otherMeasurement.finish(len(opaqueFindings))
	findings = append(findings, opaqueFindings...)
	if plan.runsProjection(procedureProjectionRuntime) {
		runtimeMeasurement := profile.begin(procedureDomainRuntime)
		runtimeFindings := a.deterministicRuntimeErrorFindingsWithArrayResult(file, proc, ctx, moduleDecls, resultStore.arrayProjection(profile))
		if enabled, known := config.AnalyzeRuleEnabled(a.Config.Analyze, "VBA249"); known && enabled {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterRuntimeCandidateProcedures)
			if proc.Graph != nil {
				profile.add(analysisstats.CounterRuntimeCFGWalks, 1)
			}
		}
		runtimeMeasurement.finish(len(runtimeFindings))
		findings = append(findings, runtimeFindings...)
	}
	if plan.runsAnyProjection(
		procedureProjectionDictionaryGuard,
		procedureProjectionDictionaryCompareMode,
		procedureProjectionDictionaryLoop,
		procedureProjectionDictionaryKeyNormalization,
		procedureProjectionDictionaryLateBound,
		procedureProjectionDictionaryCollectionMutation,
		procedureProjectionDictionaryIndexOrigin,
	) && dictionaryCollectionAnalysisEnabled(a.Config.Analyze) {
		dictionaryMeasurement := profile.begin(procedureDomainDictionary)
		dictionaryFindings, err := a.dictionaryCollectionSafetyFindings(cancelCtx, file, proc, moduleDecls)
		if err != nil {
			dictionaryMeasurement.finishOutcome(cancelCtx, 0, err)
			return nil, err
		}
		if proc.Graph != nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterDictionaryCandidateProcedures)
			profile.add(analysisstats.CounterDictionaryCFGWalks, 1)
		}
		dictionaryMeasurement.finish(len(dictionaryFindings))
		findings = append(findings, dictionaryFindings...)
	}
	if plan.runsProjection(procedureProjectionDictionaryIteration) && a.Config.Analyze.DetectDictionaryIterationValueUsage {
		dictionaryMeasurement := profile.begin(procedureDomainDictionary)
		dictionaryFindings := a.dictionaryIterationValueUsageFindings(file, proc, moduleDecls)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterDictionaryCandidateProcedures)
		dictionaryMeasurement.finish(len(dictionaryFindings))
		findings = append(findings, dictionaryFindings...)
	}

	sourceMeasurement := profile.begin(procedureDomainSourceScan)
	sourceFindingStart := len(findings)
	if profile != nil {
		start := proc.StartLine - 1
		if start < 0 {
			start = 0
		}
		end := proc.EndLine
		if end > len(file.Lines) {
			end = len(file.Lines)
		}
		if end > start {
			profile.add(analysisstats.CounterSourceLineScans, uint64(end-start))
			profile.kernel()
		}
	}
	for i := proc.StartLine - 1; i < proc.EndLine && i < len(file.Lines); i++ {
		if i&0xff == 0 {
			if err := cancelCtx.Err(); err != nil {
				sourceMeasurement.finishOutcome(cancelCtx, len(findings)-sourceFindingStart, err)
				return nil, err
			}
		}
		lineNo := i + 1
		stmt := normalizedCodeLine(file.Lines[i])
		worksheetStmt, worksheetStatementStart := worksheetLogicalStatement(file.Lines, i, proc.EndLine-1)
		if stmt == "" {
			continue
		}
		if endWithRe.MatchString(stmt) {
			if len(withStack) > 0 {
				withStack = withStack[:len(withStack)-1]
			}
			worksheetRoots.popWith()
			continue
		}
		if worksheetStatementStart && a.Config.Analyze.ForbidUnqualifiedExcelObjects {
			excelMeasurement := profile.begin(procedureDomainExcel)
			excelFindings := a.unqualifiedExcelFindings(file, proc, lineNo, normalizedCodeLine(worksheetStmt), shadowedVBA205)
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			excelMeasurement.finish(len(excelFindings))
			findings = append(findings, excelFindings...)
		}
		if m := withRe.FindStringSubmatch(stmt); len(m) > 0 {
			withStack = append(withStack, resolveWithInfo(m[1], decls))
			if worksheetStatementStart {
				if a.Config.Analyze.DetectWorksheetRootMismatch || a.Config.Analyze.DetectUnstableLastRowPatterns {
					excelMeasurement := profile.begin(procedureDomainExcel)
					excelFindings := a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)
					profile.kernel()
					profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
					excelMeasurement.finish(len(excelFindings))
					findings = append(findings, excelFindings...)
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
				excelMeasurement := profile.begin(procedureDomainExcel)
				excelFindings := a.worksheetRootFindings(file, proc, lineNo, worksheetStmt, worksheetRoots)
				profile.kernel()
				profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
				excelMeasurement.finish(len(excelFindings))
				findings = append(findings, excelFindings...)
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
			excelMeasurement := profile.begin(procedureDomainExcel)
			excelFindings := a.memberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			excelMeasurement.finish(len(excelFindings))
			findings = append(findings, excelFindings...)
		} else {
			excelMeasurement := profile.begin(procedureDomainExcel)
			excelFindings := a.legacyMemberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			excelMeasurement.finish(len(excelFindings))
			findings = append(findings, excelFindings...)
		}
		if setAssignRe.MatchString(stmt) {
			if name, receiver, ok := rangeFindAssignment(stmt, currentWithExpression(withStack)); ok {
				findAssignments[strings.ToLower(name)] = rangeFindAssignmentInfo{Line: lineNo, Receiver: receiver}
			}
			continue
		}
		if m := assignRe.FindStringSubmatch(stmt); len(m) > 0 {
			target := strings.ToLower(m[1])
			if proc.Name != "" && strings.EqualFold(target, proc.Name) && isObjectType(proc.ReturnType) {
				objectMeasurement := profile.begin(procedureDomainObject)
				objectFinding := a.objectSetFinding(file, proc, lineNo, "VBA103", m[1], proc.ReturnType)
				profile.kernel()
				profile.candidate(&candidateCounters, analysisstats.CounterObjectCandidateProcedures)
				objectMeasurement.finish(1)
				findings = append(findings, objectFinding)
				continue
			}
			if cm := callAssignRe.FindStringSubmatch(stmt); len(cm) > 0 {
				callee := strings.ToLower(lastName(cm[2]))
				if typ, ok := decls.lookup(target); ok && typ.Object && isObjectType(ctx.functionReturns[callee]) {
					objectMeasurement := profile.begin(procedureDomainObject)
					objectFinding := a.objectSetFinding(file, proc, lineNo, "VBA102", m[1], ctx.functionReturns[callee])
					profile.kernel()
					profile.candidate(&candidateCounters, analysisstats.CounterObjectCandidateProcedures)
					objectMeasurement.finish(1)
					findings = append(findings, objectFinding)
					continue
				}
			}
			if decl, ok := decls.lookup(target); ok && decl.Object {
				objectMeasurement := profile.begin(procedureDomainObject)
				objectFinding := a.objectSetFinding(file, proc, lineNo, "VBA101", m[1], decl.Type)
				profile.kernel()
				profile.candidate(&candidateCounters, analysisstats.CounterObjectCandidateProcedures)
				objectMeasurement.finish(1)
				findings = append(findings, objectFinding)
			}
		}
		if a.Config.Analyze.DetectRangeFindNothingCheck {
			excelMeasurement := profile.begin(procedureDomainExcel)
			excelFindings := a.rangeFindFindings(file, proc, lineNo, stmt, findAssignments, guardedFinds)
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			excelMeasurement.finish(len(excelFindings))
			findings = append(findings, excelFindings...)
		}
		if plan.runsProjection(procedureProjectionArrayComparison) && objectNothingEqualityLineIsExecutable(proc, lineNo, stmt) {
			arrayMeasurement := profile.begin(procedureDomainArray)
			arrayFindings := a.objectArrayComparisonFindings(file, proc, lineNo, stmt, decls)
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterArrayCandidateProcedures)
			arrayMeasurement.finish(len(arrayFindings))
			findings = append(findings, arrayFindings...)
		}
	}
	sourceMeasurement.finish(len(findings) - sourceFindingStart)
	if plan.runsProjection(procedureProjectionObject) && a.Config.Analyze.DetectObjectUseBeforeSet && ctx.objectAnalysis != nil {
		key := objectSummaryKey(file.IR.Path, objectProcedureQualifiedName(proc), string(proc.ProcedureKind), proc.StartLine)
		if objectPlan := ctx.objectAnalysis.plans[key]; objectPlan != nil {
			objectMeasurement := profile.begin(procedureDomainObject)
			objectFindings := a.objectUseBeforeSetIRFindingsPlan(objectPlan, ctx.objectAnalysis.summaries, ctx.objectAnalysis.entries[key])
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterObjectCandidateProcedures)
			objectMeasurement.finish(len(objectFindings))
			findings = append(findings, objectFindings...)
		}
	}
	if plan.runsAnyProjection(procedureProjectionArrayLifecycle, procedureProjectionArrayRedim, procedureProjectionArrayComparison) {
		arrayMeasurement := profile.begin(procedureDomainArray)
		arrayFindings := []Finding(nil)
		if result := resultStore.arrayProjection(profile); result != nil {
			arrayFindings = result.lifecycle()
		} else {
			arrayFindings = a.arrayLifecycleFindings(file, proc, ctx, moduleDecls)
		}
		if arrayResult == nil && (a.Config.Analyze.DetectArrayLifecycleSafety || a.Config.Analyze.DetectRedimPreserveDimension || a.Config.Analyze.DetectObjectArrayComparison) {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterArrayCandidateProcedures)
			if proc.Graph != nil {
				profile.add(analysisstats.CounterArrayCFGWalks, 1)
				if a.Config.Analyze.DetectArrayLifecycleSafety {
					profile.add(analysisstats.CounterArrayCFGWalks, 1)
				}
			}
		}
		arrayMeasurement.finish(len(arrayFindings))
		findings = append(findings, arrayFindings...)
		findings = suppressDeterministicArrayWarningDuplicates(findings)
	}
	if plan.runsProjection(procedureProjectionArrayRedimLoop) {
		arrayMeasurement := profile.begin(procedureDomainArray)
		redimFindings := []Finding(nil)
		if result := resultStore.arrayProjection(profile); result != nil {
			redimFindings = result.redim()
		} else {
			redimFindings = a.redimPreserveLoopFindings(file, proc, moduleDecls)
		}
		if arrayResult == nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterArrayCandidateProcedures)
			if proc.Graph != nil {
				profile.add(analysisstats.CounterArrayCFGWalks, 1)
			}
		}
		arrayMeasurement.finish(len(redimFindings))
		findings = append(findings, redimFindings...)
	}
	if plan.runsProjection(procedureProjectionApplicationRestore) && a.Config.Analyze.DetectApplicationStateRestore {
		applicationMeasurement := profile.begin(procedureDomainApplicationState)
		applicationFindings := a.applicationStateFindings(file, proc, projectEffects)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterApplicationStateCandidateProcedures)
		applicationMeasurement.finish(len(applicationFindings))
		findings = append(findings, applicationFindings...)
	}
	if plan.runsProjection(procedureProjectionApplicationEffects) && a.Config.Analyze.DetectApplicationStateCallEffects {
		applicationMeasurement := profile.begin(procedureDomainApplicationState)
		applicationFindings := a.applicationStateCallEffectFindings(file, proc, projectEffects)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterApplicationStateCandidateProcedures)
		applicationMeasurement.finish(len(applicationFindings))
		findings = append(findings, applicationFindings...)
	}
	if plan.runsProjection(procedureProjectionApplicationReentry) && a.Config.Analyze.DetectEventHandlerReentry {
		applicationMeasurement := profile.begin(procedureDomainApplicationState)
		eventFindings := a.eventHandlerReentryFindings(file, proc, projectEffects)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterApplicationStateCandidateProcedures)
		applicationMeasurement.finish(len(eventFindings))
		findings = append(findings, eventFindings...)
	}
	var plannedHTTPFindings []Finding
	if plan.runsDataflowLane(procedureDataflowLaneGeneric) || plan.runsDataflowLane(procedureDataflowLaneHTTP) {
		dataFlowFindings, httpFindings, err := a.executeDataflowLanes(cancelCtx, file, proc, plan, &resultStore, profile, &candidateCounters)
		if err != nil {
			return nil, err
		}
		findings = append(findings, dataFlowFindings...)
		plannedHTTPFindings = httpFindings
	}
	if plan.runsProjection(procedureProjectionDataflowFile) && a.Config.Analyze.DetectUnsafeFilePath {
		filePathMeasurement := profile.begin(procedureDomainDataflow)
		filePathFindings := a.filePathSafetyFindings(file, proc)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterDataflowCandidateProcedures)
		filePathMeasurement.finish(len(filePathFindings))
		findings = append(findings, filePathFindings...)
	}
	findings = append(findings, plannedHTTPFindings...)
	if plan.runsProjection(procedureProjectionResource) && a.Config.Analyze.DetectResourceLeaks {
		resourceMeasurement := profile.begin(procedureDomainResource)
		resourceFindings := a.resourceLeakFindings(file, proc)
		if proc.Graph != nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterResourceCandidateProcedures)
			profile.add(analysisstats.CounterResourceCFGWalks, 1)
		}
		resourceMeasurement.finish(len(resourceFindings))
		findings = append(findings, resourceFindings...)
	}
	if plan.runsProjection(procedureProjectionExcelLoop) && a.Config.Analyze.DetectExcelCellAccessInLoops {
		excelMeasurement := profile.begin(procedureDomainExcel)
		excelFindings := a.excelLoopAccessFindings(file, proc)
		if a.typeDB != nil && proc.Graph != nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			profile.add(analysisstats.CounterExcelCFGWalks, 1)
		}
		excelMeasurement.finish(len(excelFindings))
		findings = append(findings, excelFindings...)
	}
	if plan.runsProjection(procedureProjectionExcelInvariant) && a.Config.Analyze.DetectLoopInvariantExcelObjectResolution {
		excelMeasurement := profile.begin(procedureDomainExcel)
		excelFindings := a.excelLoopInvariantFindings(file, proc)
		if a.typeDB != nil && proc.Graph != nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
			profile.add(analysisstats.CounterExcelCFGWalks, 1)
		}
		excelMeasurement.finish(len(excelFindings))
		findings = append(findings, excelFindings...)
	}
	if plan.runsProjection(procedureProjectionExcelRange) && a.Config.Analyze.DetectExpensiveFullRangeOperations {
		excelMeasurement := profile.begin(procedureDomainExcel)
		excelFindings := a.expensiveFullRangeOperationFindings(file, proc)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
		excelMeasurement.finish(len(excelFindings))
		findings = append(findings, excelFindings...)
	}
	if plan.runsProjection(procedureProjectionExcelValue2) && a.Config.Analyze.DetectValue2PerformanceOpportunities {
		excelMeasurement := profile.begin(procedureDomainExcel)
		excelFindings := a.value2PerformanceFindings(file, proc)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterExcelCandidateProcedures)
		excelMeasurement.finish(len(excelFindings))
		findings = append(findings, excelFindings...)
	}
	if plan.runsProjection(procedureProjectionErrorHandler) && a.Config.Analyze.DetectErrorHandlerFallthrough {
		errorMeasurement := profile.begin(procedureDomainError)
		errorFindings := a.errorHandlerFallthroughFindings(file, proc)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterErrorCandidateProcedures)
		if proc.Graph != nil {
			profile.add(analysisstats.CounterErrorCFGWalks, 1)
		}
		errorMeasurement.finish(len(errorFindings))
		findings = append(findings, errorFindings...)
	}
	if plan.runsProjection(procedureProjectionErrorResume) && a.Config.Analyze.DetectLeakedOnErrorResumeNextScopes {
		errorMeasurement := profile.begin(procedureDomainError)
		errorFindings := a.leakedOnErrorResumeNextFindings(file, proc)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterErrorCandidateProcedures)
		if proc.Graph != nil {
			profile.add(analysisstats.CounterErrorCFGWalks, 1)
		}
		errorMeasurement.finish(len(errorFindings))
		findings = append(findings, errorFindings...)
	}
	if plan.runsProjection(procedureProjectionErrorSuppression) && a.Config.Analyze.DetectErrorSuppressionPropagation {
		errorMeasurement := profile.begin(procedureDomainError)
		errorFindings := a.errorSuppressionFindings(file, proc, projectEffects)
		profile.kernel()
		profile.candidate(&candidateCounters, analysisstats.CounterErrorCandidateProcedures)
		if proc.Graph != nil {
			profile.add(analysisstats.CounterErrorCFGWalks, 1)
		}
		errorMeasurement.finish(len(errorFindings))
		findings = append(findings, errorFindings...)
	}
	if plan.runsProjection(procedureProjectionArrayRangeShape) {
		arrayMeasurement := profile.begin(procedureDomainArray)
		arrayFindings := []Finding(nil)
		if result := resultStore.arrayProjection(profile); result != nil {
			arrayFindings = result.rangeShape()
		} else {
			arrayFindings = a.rangeValueShapeFindings(file, proc)
		}
		if arrayResult == nil {
			profile.kernel()
			profile.candidate(&candidateCounters, analysisstats.CounterArrayCandidateProcedures)
			if proc.Graph != nil {
				profile.add(analysisstats.CounterArrayCFGWalks, 1)
			}
		}
		arrayMeasurement.finish(len(arrayFindings))
		findings = append(findings, arrayFindings...)
	}
	otherMeasurement = profile.begin(procedureDomainOther)
	functionFindings := a.functionReturnPathFindings(file, proc)
	if a.Config.Analyze.DetectFunctionReturnPath {
		profile.kernel()
	}
	otherMeasurement.finish(len(functionFindings))
	findings = append(findings, functionFindings...)
	findings = suppressDictionaryGuardsForUninitializedObjects(findings)
	return findings, cancelCtx.Err()
}

// VBA202 is the root-cause diagnostic when a Dictionary/Collection receiver
// itself may still be Nothing.  Do not pair that finding with VBA207's key
// existence warning for the same source access; once construction is proven,
// VBA207 remains unchanged.
func suppressDictionaryGuardsForUninitializedObjects(findings []Finding) []Finding {
	unsafe := map[string]map[string]bool{}
	for _, finding := range findings {
		if finding.Code != "VBA202" {
			continue
		}
		key := finding.File + "|" + finding.Procedure + "|" + strconv.Itoa(finding.Line)
		receiver := findingReceiver(finding)
		if receiver != "" {
			if unsafe[key] == nil {
				unsafe[key] = map[string]bool{}
			}
			unsafe[key][receiver] = true
		}
	}
	if len(unsafe) == 0 {
		return findings
	}
	out := findings[:0]
	for _, finding := range findings {
		key := finding.File + "|" + finding.Procedure + "|" + strconv.Itoa(finding.Line)
		if finding.Code == "VBA207" && unsafe[key][findingReceiver(finding)] {
			continue
		}
		out = append(out, finding)
	}
	return out
}

func findingReceiver(finding Finding) string {
	message := strings.ToLower(strings.TrimSpace(finding.Message))
	var receiver string
	switch finding.Code {
	case "VBA202":
		if index := strings.Index(message, " may be "); index > 0 {
			receiver = message[:index]
		}
	case "VBA207":
		for _, marker := range []string{" accesses", " is ", " uses", " item/key"} {
			if index := strings.Index(message, marker); index > 0 {
				receiver = message[:index]
				break
			}
		}
	}
	return cleanIdentifier(strings.TrimSpace(receiver))
}

func sourceProceduresWithEffects(file parsedFile, project effects.ProjectSummary) []sourceProcedure {
	procedures := file.procedures()
	for i := range procedures {
		if i >= len(file.IR.Procedures) {
			break
		}
		id := procedureEffectIdentity(file.IR, file.IR.Procedures[i].Symbol)
		if summary, ok := project.LookupDirect(id); ok {
			procedures[i].Effects = &summary
		}
	}
	return procedures
}

func rebindRealtimeProcedureProjection(current, materialized []sourceProcedure) []sourceProcedure {
	if len(materialized) > 0 {
		return materialized
	}
	// A source file can have no recovered procedure IR. The caller installs a
	// whole-file synthetic procedure for source-line diagnostics; preserve that
	// fallback when participant materialization has nothing to rebind.
	return current
}

func (file *parsedFile) procedures() []sourceProcedure {
	return append([]sourceProcedure(nil), file.procedureProjection()...)
}

// procedureView returns an allocation-free read-only view over the cached
// procedure projection. Callers that need to retain or mutate procedure
// values must use procedures instead.
func (file *parsedFile) procedureView() readOnlySpan[sourceProcedure] {
	return newReadOnlySpan(file.procedureProjection())
}

// procedureProjection returns the cached backing projection used to construct
// owned copies and read-only views. Rule implementations must use procedures
// or procedureView instead of retaining this slice.
func (file *parsedFile) procedureProjection() []sourceProcedure {
	if file.Procedures != nil {
		return file.Procedures
	}
	return sourceProceduresFromIRRef(&file.IR, file.CFG)
}

func (file parsedFile) moduleDecls() map[string]sourceDeclaration {
	if file.ModuleDeclarations != nil {
		return file.ModuleDeclarations
	}
	if facts := file.moduleAnalysisFacts(); facts != nil {
		return facts.moduleDeclarations
	}
	return nil
}

func procedureEffectIdentity(document procedureir.DocumentIR, symbol procedureir.ProcedureSymbol) effects.ProcedureIdentity {
	return effects.ProcedureIdentity{
		File: filepath.ToSlash(filepath.Clean(document.Path)), Module: document.ModuleName, ModuleKind: document.ModuleKind,
		Name: symbol.Name, QualifiedName: symbol.QualifiedName, Kind: symbol.Kind,
		Visibility: symbol.Visibility, DeclarationLine: symbol.DeclarationRange.StartLine,
	}
}

// sourceProceduresFromIRRef materializes lightweight analyzer views over the
// canonical, Go-owned DocumentIR. It intentionally does not copy any of the
// procedure IR collections; facts and views alias the ProcedureIR storage for
// the lifetime of the document revision.
func sourceProceduresFromIRRef(document *procedureir.DocumentIR, controlFlow ...vbacfg.Document) []sourceProcedure {
	if document == nil {
		return nil
	}
	procedures := make([]sourceProcedure, 0, len(document.Procedures))
	module := strings.TrimSpace(document.ModuleName)
	if module == "" {
		module = strings.TrimSuffix(filepath.Base(document.Path), filepath.Ext(document.Path))
	}
	for procedureIndex := range document.Procedures {
		procedure := &document.Procedures[procedureIndex]
		kind := "Sub"
		switch procedure.Symbol.Kind {
		case procedureir.ProcedureFunction:
			kind = "Function"
		case procedureir.ProcedureProperty, procedureir.ProcedurePropertyGet,
			procedureir.ProcedurePropertyLet, procedureir.ProcedurePropertySet:
			kind = "Property"
		}
		source := sourceProcedure{
			Document:         document,
			IR:               procedure,
			Index:            procedureIndex,
			Kind:             kind,
			ProcedureKind:    procedure.Symbol.Kind,
			Name:             procedure.Symbol.Name,
			Module:           module,
			ModuleKind:       document.ModuleKind,
			Visibility:       procedure.Symbol.Visibility,
			ReturnType:       procedure.Symbol.ReturnType,
			ReturnValueShape: procedure.Symbol.ValueShape,
			StartLine:        procedure.Symbol.DeclarationRange.StartLine,
			EndLine:          procedure.Symbol.DeclarationRange.EndLine,
			StartByte:        procedure.Symbol.DeclarationRange.StartByte,
			EndByte:          procedure.Symbol.DeclarationRange.EndByte,
			Params:           newReadOnlySpan(procedure.Symbol.Parameters),
			Declarations:     newReadOnlySpan(procedure.Declarations),
			Statements:       newReadOnlySpan(procedure.Statements),
			Expressions:      newReadOnlySpan(procedure.Expressions),
			Calls:            newReadOnlySpan(procedure.Calls),
			Accesses:         newReadOnlySpan(procedure.Accesses),
		}
		source.Facts = newProcedureAnalysisFactsForProcedure(procedure)
		if len(controlFlow) > 0 && procedureIndex < len(controlFlow[0].Graphs) {
			source.Graph = &controlFlow[0].Graphs[procedureIndex]
		}
		graphUnknown := source.Graph != nil && len(source.Graph.UnknownFlowSources) > 0
		source.Features = finalizeProcedureFeatures(source.Facts.features, *document, *procedure, source.Graph != nil, graphUnknown)
		source.Facts.features = source.Features
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
		for _, part := range splitArgs(m[1]) {
			name, typ, array, newExpr := declarationNameAndType(part)
			if name == "" {
				continue
			}
			decls[strings.ToLower(name)] = sourceDeclaration{Name: name, Type: typ, Line: lineNo, Object: isObjectType(typ), Array: array, NewExpression: newExpr, Static: strings.HasPrefix(lower, "static ")}
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
	facts := buildModuleAnalysisFacts(lines, procedureir.DocumentIR{}, procedures)
	return facts.moduleDeclarations
}

// lineInAnyProcedure is retained for package-local compatibility. Normal
// analysis uses moduleAnalysisFacts.procedureLineOwners; this helper handles
// the already-sorted procedure projection with a binary search and preserves
// the old behavior for synthetic unsorted callers.
func lineInAnyProcedure(line int, procedures []sourceProcedure) bool {
	if len(procedures) == 0 {
		return false
	}
	sorted := true
	for index := 1; index < len(procedures); index++ {
		if procedures[index].StartLine < procedures[index-1].StartLine {
			sorted = false
			break
		}
	}
	if !sorted {
		for _, procedure := range procedures {
			if procedure.StartLine <= line && line <= procedure.EndLine {
				return true
			}
		}
		return false
	}
	index := sort.Search(len(procedures), func(index int) bool { return procedures[index].StartLine > line })
	if index == 0 {
		return false
	}
	procedure := procedures[index-1]
	return procedure.StartLine <= line && line <= procedure.EndLine
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

func (a Analyzer) legacyMemberMismatchFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls declarationScope, withStack []withInfo) []Finding {
	all := a.memberMismatchFindings(file, proc, lineNo, stmt, decls, withStack)
	filtered := all[:0]
	for _, finding := range all {
		if finding.Code == "VBA104" {
			filtered = append(filtered, finding)
		}
	}
	return filtered
}

func (a Analyzer) memberMismatchFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls declarationScope, withStack []withInfo) []Finding {
	var findings []Finding
	if currentWith, ok := currentWithInfo(withStack); ok {
		if m := withMemberRe.FindStringSubmatch(stmt); len(m) > 0 {
			if rule, ok := invalidMemberRuleFor(currentWith.Type, m[1]); ok {
				findings = append(findings, a.memberFinding(file, proc, lineNo, currentWith.Target, currentWith.Type, m[1], rule))
			}
		}
	}
	for _, m := range memberRe.FindAllStringSubmatch(stmt, -1) {
		if decl, ok := decls.lookup(m[1]); ok {
			if rule, ok := invalidMemberRuleFor(decl.Type, m[2]); ok {
				findings = append(findings, a.memberFinding(file, proc, lineNo, m[1], decl.Type, m[2], rule))
			}
		}
	}
	return findings
}

type rangeFindAssignmentInfo struct {
	Line     int
	Receiver string
}

func (a Analyzer) rangeFindFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, findAssignments map[string]rangeFindAssignmentInfo, guarded map[string]bool) []Finding {
	lower := strings.ToLower(stmt)
	for name := range findAssignments {
		if strings.Contains(lower, "if "+name+" is nothing") || strings.Contains(lower, "if not "+name+" is nothing") {
			guarded[name] = true
		}
	}
	if name, receiver, ok := rangeFindAssignment(stmt, ""); ok {
		findAssignments[strings.ToLower(name)] = rangeFindAssignmentInfo{Line: lineNo, Receiver: receiver}
		return nil
	}
	var findings []Finding
	for name, assignment := range findAssignments {
		if guarded[name] {
			continue
		}
		if !a.rangeFindReceiverIsExcelRange(file, assignment.Receiver, assignment.Line) {
			continue
		}
		if strings.Contains(lower, name+".") {
			suggestion := "Add If " + name + " Is Nothing Then handling after the Find assignment."
			if assignment.Line == 0 {
				suggestion = "Check the Find result for Nothing before dereferencing it."
			}
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA201", "warning", "Range.Find result "+name+" is dereferenced before a Nothing check.", "Range.Find returns Nothing when no match is found, so dereferencing the result can raise runtime error 91.", suggestion))
			guarded[name] = true
		}
	}
	return findings
}

func (a Analyzer) unqualifiedExcelFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, shadowed vba205ShadowedIdentifiers) []Finding {
	if vba205NonExecutableStatement(stmt) {
		return nil
	}
	var findings []Finding
	for _, m := range unqualifiedExcelRe.FindAllStringSubmatchIndex(stmt, -1) {
		name := stmt[m[4]:m[5]]
		if vba205QualifiedOrShadowedRoot(stmt[m[0]:m[1]], name, shadowed) {
			continue
		}
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Unqualified "+name+" access depends on the active worksheet.", "Unqualified Excel object access is resolved through the active sheet or selection at runtime.", "Qualify "+name+" with an explicit Worksheet or Range object."))
	}
	for _, m := range activeExcelRe.FindAllStringSubmatchIndex(stmt, -1) {
		name := stmt[m[2]:m[3]]
		if vba205QualifiedOrShadowedRoot(stmt[m[0]:m[1]], name, shadowed) {
			continue
		}
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", name+" creates an active Excel object dependency.", name+" depends on the user's current Excel UI state and can target a different object during automation.", "Pass an explicit Workbook, Worksheet, or Range argument instead."))
	}
	for _, m := range unqualifiedSheetCollectionRe.FindAllStringSubmatchIndex(stmt, -1) {
		name := stmt[m[2]:m[3]]
		if vba205QualifiedOrShadowedRoot(stmt[m[0]:m[1]], name, shadowed) {
			continue
		}
		suggestion := "Use ThisWorkbook." + name + "(...) or select " + name + " from an explicit Workbook argument."
		if addInStandardModule(a.Config, file) {
			suggestion = "Select " + name + " from an explicit caller Workbook argument."
		}
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Unqualified "+name+" access depends on the active workbook.", "The unqualified "+name+" collection is resolved from Excel's active workbook at runtime.", suggestion))
	}
	for _, m := range positionalExcelCollectionRe.FindAllStringSubmatchIndex(stmt, -1) {
		name := stmt[m[2]:m[3]]
		if vba205QualifiedOrShadowedRoot(stmt[m[0]:m[1]], name, shadowed) {
			continue
		}
		index := stmt[m[4]:m[5]]
		root := name + "(" + index + ")"
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", root+" depends on Excel collection ordering.", root+" can select a different object when workbook or window order changes.", "Select the target by name or receive an explicit "+strings.TrimSuffix(name, "s")+" argument."))
	}
	for _, open := range workbooksOpenRe.FindAllStringSubmatchIndex(stmt, -1) {
		if vba205QualifiedOrShadowedRoot(stmt[open[0]:open[1]], stmt[open[2]:open[3]], shadowed) {
			continue
		}
		if !capturedWorkbooksOpen(stmt, open[2]) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "Workbooks.Open result is not captured.", "An uncaptured Workbooks.Open result forces later code to depend on active workbook state.", "Capture the opened workbook: Set wb = Workbooks.Open(...)."))
		}
	}
	if addInStandardModule(a.Config, file) && thisWorkbookRe.MatchString(stmt) {
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA205", "warning", "ThisWorkbook in an add-in targets the add-in workbook.", "Inside an add-in standard module, ThisWorkbook is the add-in rather than the caller workbook.", "Receive the caller workbook as an explicit Workbook argument."))
	}
	return findings
}

func vba205NonExecutableStatement(stmt string) bool {
	lower := strings.ToLower(strings.TrimSpace(stmt))
	return isProcedureHeaderLine(lower) ||
		strings.HasPrefix(lower, "attribute ") ||
		strings.HasPrefix(lower, "option ") ||
		strings.HasPrefix(lower, "implements ") ||
		strings.HasPrefix(lower, "declare ")
}

type vba205ShadowedIdentifiers struct {
	decls      declarationScope
	local      map[string]struct{}
	procedures map[string]procedureSignature
	resolver   procedureir.Resolver
	proc       sourceProcedure
	facts      *moduleAnalysisFacts
}

func vba205ShadowedIdentifiersWithFacts(proc sourceProcedure, decls declarationScope, ctx analysisContext, facts *moduleAnalysisFacts) vba205ShadowedIdentifiers {
	shadowed := vba205ShadowedIdentifiers{
		decls:      decls,
		local:      make(map[string]struct{}, proc.Accesses.Len()+proc.Declarations.Len()),
		procedures: ctx.procedures,
		resolver:   ctx.procedureResolver,
		proc:       proc,
		facts:      facts,
	}
	for declaration := range proc.Declarations.All() {
		switch declaration.Scope {
		case procedureir.ScopeParameter, procedureir.ScopeLocal, procedureir.ScopeModule, procedureir.ScopeProject:
			shadowed.local[strings.ToLower(declaration.Name)] = struct{}{}
		}
	}
	for access := range proc.Accesses.All() {
		switch access.Scope {
		case procedureir.ScopeParameter, procedureir.ScopeLocal, procedureir.ScopeModule, procedureir.ScopeProject:
			shadowed.local[strings.ToLower(access.Name)] = struct{}{}
		}
	}
	return shadowed
}

func (shadowed vba205ShadowedIdentifiers) contains(name string) bool {
	key := strings.ToLower(strings.TrimSpace(name))
	if key == "" {
		return false
	}
	if _, ok := shadowed.local[key]; ok {
		return true
	}
	if _, ok := shadowed.decls.lookup(key); ok {
		return true
	}
	// A procedure declared in the caller's own module is always visible to
	// unqualified VBA calls (including Private procedures). Resolve only the
	// candidate root seen by the source scanner instead of copying every
	// project procedure into every procedure's shadow set.
	if shadowed.facts != nil && shadowed.facts.hasProcedure(key) {
		return true
	}
	if _, ok := shadowed.procedures[key]; ok {
		return vba205ProcedureVisibleFrom(shadowed.resolver, shadowed.proc, key)
	}
	return false
}

func vba205ProcedureVisibleFrom(resolver procedureir.Resolver, proc sourceProcedure, name string) bool {
	caller := procedureir.ProcedureRef{
		Name: proc.Name, Kind: proc.ProcedureKind, QualifiedName: proc.Module + "." + proc.Name,
	}
	resolution := resolver.ResolveCall(procedureir.CallSite{
		Caller: caller,
		Callee: procedureir.Callee{Text: name, BaseName: name, Member: name},
	})
	return resolution.Status == procedureir.ResolutionMatched || resolution.Status == procedureir.ResolutionAmbiguous
}

func vba205QualifiedOrShadowedRoot(match, name string, shadowed vba205ShadowedIdentifiers) bool {
	lowerMatch := strings.ToLower(match)
	if strings.Contains(lowerMatch, "application.") {
		return false
	}
	rootName := strings.ToLower(name)
	if dot := strings.IndexByte(rootName, '.'); dot >= 0 {
		rootName = rootName[:dot]
	}
	return shadowed.contains(rootName)
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

type dictionaryIterationLoop struct {
	item     string
	dict     string
	reported bool
}

// dictionaryIterationValueUsageFindings reports only direct dictionary
// iteration whose control variable is subsequently used as an object. VBA
// iterates Dictionary keys, so ordinary key usage must remain silent.
func (a Analyzer) dictionaryIterationValueUsageFindings(file parsedFile, proc sourceProcedure, moduleDecls map[string]sourceDeclaration) []Finding {
	decls := newDeclarationScope(file, proc)
	decls.module = moduleDecls

	declaredDictionaries := map[string]bool{}
	decls.forEach(func(key string, decl sourceDeclaration) {
		if isDictionaryType(decl.Type) {
			declaredDictionaries[key] = true
		}
	})
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
			} else if source := strings.ToLower(cleanIdentifier(rhs)); dcBareIdentifier(rhs) && (declaredDictionaries[source] || inferredDictionaries[source]) {
				inferredDictionaries[target] = true
			} else if helper, _, called := dcSimpleFunctionCall(rhs); called {
				if summary, ok := a.dcHelper(helper); ok && summary.Factory == dcDictionary {
					inferredDictionaries[target] = true
				} else {
					delete(inferredDictionaries, target)
				}
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

func dictionaryIterationValueUse(stmt, item string, decls declarationScope) bool {
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
	target, ok := decls.lookup(match[1])
	if !ok || !target.Object {
		return false
	}
	rhs := strings.TrimSpace(stmt[strings.Index(stmt, "=")+1:])
	return strings.EqualFold(cleanIdentifier(rhs), itemKey)
}

func (a Analyzer) objectArrayComparisonFindings(file parsedFile, proc sourceProcedure, lineNo int, stmt string, decls declarationScope) []Finding {
	var findings []Finding
	reported := make(map[string]bool)
	decls.forEach(func(key string, decl sourceDeclaration) {
		if decl.Object && objectNothingEqualityComparisonExists(stmt, key) {
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA209", "warning", decl.Name+" is compared to Nothing with =.", "Object references must be compared with Is Nothing, not the scalar equality operator.", "Use `If "+decl.Name+" Is Nothing Then` or `If Not "+decl.Name+" Is Nothing Then`."))
			reported[key] = true
		}
	})
	for declaration := range proc.Declarations.All() {
		if declaration.Scope != procedureir.ScopeParameter || !sourceDeclarationIsObject(declaration, declaration.Type) {
			continue
		}
		key := strings.ToLower(strings.TrimSpace(declaration.Name))
		if reported[key] || !objectNothingEqualityComparisonExists(stmt, key) {
			continue
		}
		findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA209", "warning", declaration.Name+" is compared to Nothing with =.", "Object references must be compared with Is Nothing, not the scalar equality operator.", "Use `If "+declaration.Name+" Is Nothing Then` or `If Not "+declaration.Name+" Is Nothing Then`."))
		reported[key] = true
	}
	return findings
}

func objectNothingEqualityLineIsExecutable(proc sourceProcedure, lineNo int, stmt string) bool {
	lower := strings.ToLower(strings.TrimSpace(stmt))
	if lineNo == proc.StartLine && isProcedureHeaderLine(lower) {
		return false
	}
	for parameter := range proc.Params.All() {
		if !parameter.Optional || parameter.Range.StartLine == 0 || lineNo < parameter.Range.StartLine || lineNo > parameter.Range.EndLine {
			continue
		}
		// Optional defaults belong to the procedure declaration, not to an
		// executable object comparison. Keep the later `If ad = Nothing` line
		// eligible by limiting this to the parameter's source range.
		return false
	}
	return true
}

func objectNothingEqualityComparisonExists(stmt, name string) bool {
	lower := strings.ToLower(stmt)
	needle := strings.ToLower(name) + " = nothing"
	for offset := 0; offset < len(lower); {
		relative := strings.Index(lower[offset:], needle)
		if relative < 0 {
			return false
		}
		start := offset + relative
		prefix := strings.TrimRightFunc(lower[:start], unicode.IsSpace)
		if !strings.HasSuffix(prefix, "set") {
			return true
		}
		setStart := len(prefix) - len("set")
		if setStart == 0 || !isVBAIdentifierRune(rune(prefix[setStart-1])) {
			offset = start + len(needle)
			continue
		}
		return true
	}
	return false
}

func (a Analyzer) applicationStateFindings(file parsedFile, proc sourceProcedure, project effects.ProjectSummary) []Finding {
	var findings []Finding
	var origins []applicationStateLeakOrigin
	if a.applicationStateLeaks != nil {
		origins = a.applicationStateLeaks.originsFor(proc)
	} else {
		origins = applicationStateLeakOrigins(proc, project)
	}
	for _, origin := range origins {
		property := applicationStatePropertyName(origin.Property)
		findings = append(findings, a.simpleFinding(
			file, proc, origin.Line, "VBA203", "warning",
			"Application."+property+" can reach "+origin.Witness.Kind+" without restoring its previous value.",
			"The changed Application."+property+" value can leave this procedure through "+origin.Witness.description()+".",
			"Save the previous Application."+property+" value and restore it in a cleanup path.",
		))
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
	Dirty     map[int]bool
	Restores  map[int]bool
	GuardedBy map[int]bool
	Unknown   bool
}

type applicationStateFlow struct {
	Dirty              map[int]bool
	Saved              map[string]applicationStateSnapshot
	ViaExceptional     bool
	dirtyShared        bool
	savedShared        bool
	snapshotMapsShared bool
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
func applicationStateExitWitnesses(proc sourceProcedure, property string, facts *procedureAnalysisFacts) map[int]applicationStateExitWitness {
	graph := proc.Graph.WithoutNormalErrRaiseContinuationView()
	in := map[vbacfg.BlockID]applicationStateFlow{graph.Entry(): newApplicationStateFlow()}
	queued := map[vbacfg.BlockID]bool{graph.Entry(): true}
	queue := []vbacfg.BlockID{graph.Entry()}

	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		queued[blockID] = false
		state := in[blockID]
		block, ok := applicationStateBlock(graph, blockID)
		if !ok {
			continue
		}
		graph.ForEachOutgoing(blockID, func(edge vbacfg.Edge) bool {
			out := applicationStateFlowAcrossEdge(proc, state, block, edge, property, facts)
			merged, changed := mergeApplicationStateFlow(in[edge.To], out)
			if !changed {
				return true
			}
			in[edge.To] = merged
			if !queued[edge.To] {
				queue = append(queue, edge.To)
				queued[edge.To] = true
			}
			return true
		})
	}
	witnesses := map[int]applicationStateExitWitness{}
	graph.ForEachEdge(func(edge vbacfg.Edge) bool {
		kind := applicationStateViewExitKind(graph, edge.To)
		if kind == "" {
			return true
		}
		state, reachable := in[edge.From]
		if !reachable || state.Dirty == nil {
			return true
		}
		block, ok := applicationStateBlock(graph, edge.From)
		if !ok {
			return true
		}
		out := applicationStateFlowAcrossEdge(proc, state, block, edge, property, facts)
		if out.ViaExceptional && kind == "normal exit" {
			kind = "error-handler path to normal exit"
		}
		for origin := range out.Dirty {
			if existing, exists := witnesses[origin]; !exists || applicationStateWitnessRank(kind) > applicationStateWitnessRank(existing.Kind) {
				witnesses[origin] = applicationStateExitWitness{Kind: kind, Line: edge.Range.StartLine}
			}
		}
		return true
	})
	return witnesses
}

func applicationStateFlowAcrossEdge(proc sourceProcedure, state applicationStateFlow, block vbacfg.Block, edge vbacfg.Edge, property string, facts *procedureAnalysisFacts) applicationStateFlow {
	out := cloneApplicationStateFlow(state)
	out.ViaExceptional = out.ViaExceptional || edge.Class == vbacfg.EdgeExceptional
	if block.Statement == nil {
		return out
	}
	switch {
	case edge.Class == vbacfg.EdgeNormal || applicationStateSavedRestore(proc, state, *block.Statement, property, facts):
		return applyApplicationStateStatement(proc, out, *block.Statement, property, facts)
	case edge.Class == vbacfg.EdgeExceptional:
		return applyApplicationStateExceptionalRestore(proc, out, *block.Statement, property, facts)
	default:
		return out
	}
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

func applicationStateBlock(graph vbacfg.CFGView, id vbacfg.BlockID) (vbacfg.Block, bool) {
	return graph.BlockByID(id)
}

func applicationStateViewExitKind(graph vbacfg.CFGView, id vbacfg.BlockID) string {
	switch id {
	case graph.NormalExit():
		return "normal exit"
	case graph.ExceptionalExit():
		return "error exit"
	case graph.TerminationExit():
		return "termination exit"
	case graph.UnknownExit():
		return "unknown exit"
	default:
		return ""
	}
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
	// CFG edges fork this state frequently. Share the map headers and defer
	// copying until a branch actually changes one of the maps. Snapshot maps
	// are treated as shared as well because Saved values can alias one another.
	out := applicationStateFlow{
		Dirty:              in.Dirty,
		Saved:              in.Saved,
		ViaExceptional:     in.ViaExceptional,
		dirtyShared:        true,
		savedShared:        true,
		snapshotMapsShared: true,
	}
	if out.Dirty == nil {
		out.Dirty = map[int]bool{}
		out.dirtyShared = false
	}
	if out.Saved == nil {
		out.Saved = map[string]applicationStateSnapshot{}
		out.savedShared = false
	}
	return out
}

func (s *applicationStateFlow) ensureDirty() {
	if s == nil || !s.dirtyShared {
		return
	}
	out := make(map[int]bool, len(s.Dirty))
	for origin := range s.Dirty {
		out[origin] = true
	}
	s.Dirty = out
	s.dirtyShared = false
}

func (s *applicationStateFlow) ensureSaved() {
	if s == nil || !s.savedShared {
		return
	}
	out := make(map[string]applicationStateSnapshot, len(s.Saved))
	for key, snapshot := range s.Saved {
		out[key] = snapshot
	}
	s.Saved = out
	s.savedShared = false
}

func (s *applicationStateFlow) mutableSavedSnapshot(key string) (applicationStateSnapshot, bool) {
	if s == nil {
		return applicationStateSnapshot{}, false
	}
	s.ensureSaved()
	snapshot, exists := s.Saved[key]
	if !exists {
		return applicationStateSnapshot{}, false
	}
	if s.snapshotMapsShared {
		snapshot.Dirty = cloneIntBoolSet(snapshot.Dirty)
		snapshot.Restores = cloneIntBoolSet(snapshot.Restores)
		snapshot.GuardedBy = cloneIntBoolSet(snapshot.GuardedBy)
	}
	return snapshot, true
}

func cloneIntBoolSet(in map[int]bool) map[int]bool {
	out := make(map[int]bool, len(in))
	for key := range in {
		out[key] = true
	}
	return out
}

func cloneApplicationStateSnapshot(in applicationStateSnapshot) applicationStateSnapshot {
	out := applicationStateSnapshot{
		Dirty:     map[int]bool{},
		Restores:  map[int]bool{},
		GuardedBy: map[int]bool{},
		Unknown:   in.Unknown,
	}
	for origin := range in.Dirty {
		out.Dirty[origin] = true
	}
	for origin := range in.Restores {
		out.Restores[origin] = true
	}
	for statementID := range in.GuardedBy {
		out.GuardedBy[statementID] = true
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
			next.ensureDirty()
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
			left = applicationStateSnapshot{Dirty: map[int]bool{}, Restores: map[int]bool{}, GuardedBy: map[int]bool{}, Unknown: true}
		}
		if !rightOK {
			right = applicationStateSnapshot{Dirty: map[int]bool{}, Restores: map[int]bool{}, GuardedBy: map[int]bool{}, Unknown: true}
		}
		merged := cloneApplicationStateSnapshot(left)
		merged.Unknown = left.Unknown || right.Unknown
		for origin := range right.Dirty {
			merged.Dirty[origin] = true
		}
		for statementID := range right.GuardedBy {
			merged.GuardedBy[statementID] = true
		}
		restoreCandidates := map[int]bool{}
		for origin := range left.Restores {
			restoreCandidates[origin] = true
		}
		for origin := range right.Restores {
			restoreCandidates[origin] = true
		}
		merged.Restores = map[int]bool{}
		for origin := range restoreCandidates {
			leftCovered := !current.Dirty[origin] || left.Restores[origin]
			rightCovered := !incoming.Dirty[origin] || right.Restores[origin]
			if leftCovered && rightCovered {
				merged.Restores[origin] = true
			}
		}
		if !applicationStateSnapshotEqual(current.Saved[key], merged) {
			next.ensureSaved()
			next.Saved[key] = merged
			changed = true
		}
	}
	return next, changed
}

func applicationStateSnapshotEqual(a, b applicationStateSnapshot) bool {
	if a.Unknown != b.Unknown || len(a.Dirty) != len(b.Dirty) || len(a.Restores) != len(b.Restores) || len(a.GuardedBy) != len(b.GuardedBy) {
		return false
	}
	for origin := range a.Dirty {
		if !b.Dirty[origin] {
			return false
		}
	}
	for origin := range a.Restores {
		if !b.Restores[origin] {
			return false
		}
	}
	for statementID := range a.GuardedBy {
		if !b.GuardedBy[statementID] {
			return false
		}
	}
	return true
}

func applyApplicationStateStatement(proc sourceProcedure, state applicationStateFlow, statement procedureir.Statement, property string, facts *procedureAnalysisFacts) applicationStateFlow {
	if statement.Recovered || statement.Kind != procedureir.StatementAssignment || statement.Target == nil {
		return state
	}
	if assignedProperty, value, isPropertyWrite := applicationPropertyAssignment(statement, facts); isPropertyWrite && assignedProperty == property {
		if applicationStateSavedRestore(proc, state, statement, property, facts) {
			variable, _ := applicationStateVariable(proc, statement.ID, value, procedureir.AccessRead)
			state.Dirty = cloneIntBoolSet(state.Saved[variable].Dirty)
			state.dirtyShared = false
			return state
		}
		if variable, ok := applicationStateVariable(proc, statement.ID, value, procedureir.AccessRead); ok {
			if saved, exists := state.Saved[variable]; exists {
				if len(saved.Restores) > 0 {
					state.ensureDirty()
				}
				for origin := range saved.Restores {
					delete(state.Dirty, origin)
				}
				if applicationStateMatchingGuard(proc, saved, statement, facts) {
					state.ensureDirty()
					for origin := range saved.Dirty {
						state.Dirty[origin] = true
					}
					return state
				}
			}
		}
		// A fully known restore variable returned above. The remaining path can
		// still carry merged per-origin coverage in an unknown snapshot.
		for key, saved := range state.Saved {
			// Loop backedges can revisit this assignment after its origin was saved;
			// an origin must never be recorded as restoring itself.
			if !saved.Unknown && !saved.Dirty[statement.ID] {
				saved, _ = state.mutableSavedSnapshot(key)
				if saved.Restores == nil {
					saved.Restores = map[int]bool{}
				}
				saved.Restores[statement.ID] = true
				state.ensureSaved()
				state.Saved[key] = saved
			}
		}
		state.ensureDirty()
		state.Dirty[statement.ID] = true
		return state
	}
	variable, ok := applicationStateVariable(proc, statement.ID, statement.Target.Text, procedureir.AccessWrite)
	if !ok {
		return state
	}
	if isApplicationPropertyReference(statement.Value, statement, facts, property) {
		state.ensureSaved()
		state.Saved[variable] = applicationStateSnapshot{
			Dirty:     cloneIntBoolSet(state.Dirty),
			Restores:  map[int]bool{},
			GuardedBy: applicationStateDirectGuard(statement, facts),
		}
		return state
	}
	if source, ok := applicationStateVariable(proc, statement.ID, expressionText(statement.Value), procedureir.AccessRead); ok {
		if saved, exists := state.Saved[source]; exists {
			state.ensureSaved()
			state.Saved[variable] = saved
			state.snapshotMapsShared = true
			return state
		}
	}
	state.ensureSaved()
	state.Saved[variable] = applicationStateSnapshot{Dirty: map[int]bool{}, Restores: map[int]bool{}, GuardedBy: map[int]bool{}, Unknown: true}
	return state
}

func applicationStateDirectGuard(statement procedureir.Statement, facts *procedureAnalysisFacts) map[int]bool {
	guards := map[int]bool{}
	visited := map[int]bool{}
	for parentID := statement.ParentID; parentID != 0; {
		if visited[parentID] {
			break
		}
		visited[parentID] = true
		parent, ok := facts.Statement(parentID)
		if !ok {
			break
		}
		if parent.Kind == procedureir.StatementElse {
			break
		}
		if parent.Kind == procedureir.StatementIf && parent.Condition != nil {
			guards[parent.ID] = true
			break
		}
		parentID = parent.ParentID
	}
	return guards
}

func applyApplicationStateExceptionalRestore(proc sourceProcedure, state applicationStateFlow, statement procedureir.Statement, property string, facts *procedureAnalysisFacts) applicationStateFlow {
	assignedProperty, value, isPropertyWrite := applicationPropertyAssignment(statement, facts)
	if !isPropertyWrite || assignedProperty != property {
		return state
	}
	variable, ok := applicationStateVariable(proc, statement.ID, value, procedureir.AccessRead)
	if !ok {
		return state
	}
	saved, exists := state.Saved[variable]
	if !exists {
		return state
	}
	if applicationStateMatchingGuard(proc, saved, statement, facts) {
		state.Dirty = cloneIntBoolSet(saved.Dirty)
		state.dirtyShared = false
		return state
	}
	if len(saved.Restores) > 0 {
		state.ensureDirty()
	}
	for origin := range saved.Restores {
		delete(state.Dirty, origin)
	}
	return state
}

func applicationStateMatchingGuard(proc sourceProcedure, saved applicationStateSnapshot, restore procedureir.Statement, facts *procedureAnalysisFacts) bool {
	restoreGuards := applicationStateDirectGuard(restore, facts)
	for savedID := range saved.GuardedBy {
		savedGuard, savedOK := facts.Statement(savedID)
		if !savedOK || savedGuard.Condition == nil {
			continue
		}
		for restoreID := range restoreGuards {
			restoreGuard, restoreOK := facts.Statement(restoreID)
			if !restoreOK || restoreGuard.Condition == nil ||
				applicationStateGuardConditionKey(savedGuard.Condition.Text) != applicationStateGuardConditionKey(restoreGuard.Condition.Text) {
				continue
			}
			if applicationStateGuardBindingsStable(proc, savedGuard, restoreGuard) {
				return true
			}
		}
	}
	return false
}

func applicationStateGuardConditionKey(condition string) string {
	runes := []rune(condition)
	var key strings.Builder
	inString := false
	for i := 0; i < len(runes); i++ {
		r := runes[i]
		if inString {
			key.WriteRune(r)
			if r != '"' {
				continue
			}
			if i+1 < len(runes) && runes[i+1] == '"' {
				key.WriteRune(runes[i+1])
				i++
				continue
			}
			inString = false
			continue
		}
		if r == '"' {
			inString = true
			key.WriteRune(r)
			continue
		}
		if unicode.IsSpace(r) {
			continue
		}
		key.WriteRune(unicode.ToLower(r))
	}
	return key.String()
}

func applicationStateGuardBindingsStable(proc sourceProcedure, savedGuard, restoreGuard procedureir.Statement) bool {
	type binding struct {
		scope procedureir.SymbolScope
		name  string
	}
	savedBindings := map[binding]bool{}
	restoreBindings := map[binding]bool{}
	for access := range proc.Accesses.All() {
		if access.Mode != procedureir.AccessRead || access.Scope == procedureir.ScopeUnresolved {
			continue
		}
		key := binding{scope: access.Scope, name: strings.ToLower(access.Name)}
		switch access.StatementID {
		case savedGuard.ID:
			savedBindings[key] = true
		case restoreGuard.ID:
			restoreBindings[key] = true
		}
	}
	if len(savedBindings) == 0 || len(savedBindings) != len(restoreBindings) {
		return false
	}
	for key := range savedBindings {
		if !restoreBindings[key] {
			return false
		}
	}
	for key := range savedBindings {
		if key.scope != procedureir.ScopeModule && key.scope != procedureir.ScopeProject {
			continue
		}
		for call := range proc.Calls.All() {
			if call.StatementID > savedGuard.ID && call.StatementID < restoreGuard.ID {
				return false
			}
		}
	}
	for access := range proc.Accesses.All() {
		if access.StatementID <= savedGuard.ID || access.StatementID >= restoreGuard.ID ||
			(access.Mode != procedureir.AccessWrite && access.Mode != procedureir.AccessReadWrite) {
			continue
		}
		if savedBindings[binding{scope: access.Scope, name: strings.ToLower(access.Name)}] {
			return false
		}
	}
	return true
}

func applicationStateSavedRestore(proc sourceProcedure, state applicationStateFlow, statement procedureir.Statement, property string, facts *procedureAnalysisFacts) bool {
	assignedProperty, value, isPropertyWrite := applicationPropertyAssignment(statement, facts)
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
	for access := range proc.Accesses.All() {
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

func applicationPropertyAssignment(statement procedureir.Statement, facts *procedureAnalysisFacts) (string, string, bool) {
	if statement.Target == nil || statement.Value == nil {
		return "", "", false
	}
	property, ok := applicationPropertyTarget(statement.Target.Text, statement, facts)
	if !ok {
		return "", "", false
	}
	return property, statement.Value.Text, true
}

func isApplicationPropertyReference(expr *procedureir.Expression, statement procedureir.Statement, facts *procedureAnalysisFacts, property string) bool {
	if expr == nil {
		return false
	}
	got, ok := applicationPropertyTarget(expr.Text, statement, facts)
	return ok && got == property
}

func applicationPropertyTarget(expression string, statement procedureir.Statement, facts *procedureAnalysisFacts) (string, bool) {
	compact := strings.ToLower(compactStatement(expression))
	for _, property := range applicationStateProperties() {
		if compact == "application."+property.Key {
			return property.Key, true
		}
		if compact == "."+property.Key && statementWithinApplicationWith(statement, facts) {
			return property.Key, true
		}
	}
	return "", false
}

func statementWithinApplicationWith(statement procedureir.Statement, facts *procedureAnalysisFacts) bool {
	visited := map[int]bool{}
	for parentID := statement.ParentID; parentID != 0; {
		if visited[parentID] {
			return false
		}
		visited[parentID] = true
		parent, ok := facts.Statement(parentID)
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
		for _, direct := range project.AllDirect() {
			if !pairNames[strings.ToLower(direct.Identity.Name)] {
				continue
			}
			if strings.EqualFold(direct.Identity.Module, proc.Effects.Identity.Module) {
				if summary, ok := project.Lookup(direct.Identity); ok {
					sameModule = append(sameModule, summary)
				}
				continue
			}
			if isProjectVisibleProcedure(direct.Identity) {
				if summary, ok := project.Lookup(direct.Identity); ok {
					projectVisible = append(projectVisible, summary)
				}
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
	for _, direct := range project.AllDirect() {
		if !direct.Has(effects.RestoresApplicationState) {
			continue
		}
		candidate, ok := project.Lookup(direct.Identity)
		if !ok || (!hasApplicationStateEffectFrom(candidate.Direct, effects.RestoresApplicationState, prop, proc.Effects.Identity) &&
			!hasApplicationStateEffectFrom(candidate.Propagated, effects.RestoresApplicationState, prop, proc.Effects.Identity)) {
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
	for _, push := range project.AllDirect() {
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
	for statement := range proc.Statements.All() {
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
		flowGraph := proc.Graph.WithoutNormalErrRaiseContinuationView()
		reachable := map[vbacfg.BlockID]bool{}
		for _, blockID := range proc.Graph.Reachable(vbacfg.EdgeFilter{NormalOnly: true}) {
			reachable[blockID] = true
		}
		for statement := range proc.Statements.All() {
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
			loopExitVariables := map[string]bool{}
			flowGraph.ForEachIncoming(block.ID, func(edge vbacfg.Edge) bool {
				if edge.Class == vbacfg.EdgeNormal && reachable[edge.From] &&
					edge.Kind != vbacfg.EdgeGoto && edge.Kind != vbacfg.EdgeUnknown {
					implicitEntry = true
					if edge.Kind == vbacfg.EdgeLoopExit {
						for candidate := range proc.Statements.All() {
							if candidate.ID == edge.StatementID && candidate.Control != nil && candidate.Control.LoopVariable != "" {
								loopExitVariables[strings.ToLower(candidate.Control.LoopVariable)] = true
							}
						}
					}
				}
				return true
			})
			if !implicitEntry {
				continue
			}
			if isSemanticCleanupFallthrough(proc, block, flowGraph, loopExitVariables) {
				continue
			}
			lineNo := statement.Range.StartLine
			findings = append(findings, a.simpleFinding(file, proc, lineNo, "VBA204", "warning", "Normal execution can fall through into error handler "+label+".", "Without Exit Sub, Exit Function, or Exit Property before the handler label, successful execution can run error handling code.", errorHandlerFallthroughSuggestion(proc, label)))
		}
		return findings
	}

	lastCodeByParent := map[int]string{}
	for statement := range proc.Statements.All() {
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
	for statement := range proc.Statements.All() {
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
		finding := a.simpleFinding(
			file, proc, outcome.startLine, "VBA214", "warning",
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

func resumeNextScopeCallRisk(calls readOnlySpan[procedureir.CallSite], statementID int) (call bool, projectCall bool) {
	for candidate := range calls.All() {
		if candidate.StatementID != statementID {
			continue
		}
		switch candidate.Resolution.Status {
		case procedureir.ResolutionMatched:
			if len(candidate.Resolution.Candidates) == 1 {
				return true, true
			}
		case procedureir.ResolutionAmbiguous, procedureir.ResolutionUnresolved, procedureir.ResolutionNotAttempted,
			procedureir.ResolutionDynamic, procedureir.ResolutionIncomplete, procedureir.ResolutionNonCallable:
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
	blocking, _ := PartitionPreflightFindings(findings, staticpreflight.NewPolicy(nil))
	return blocking
}

func PartitionPreflightFindings(findings []Finding, policy staticpreflight.Policy) (blocking []Finding, allowed []Finding) {
	return staticpreflight.Partition(findings, func(finding Finding) string { return finding.Code }, policy)
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

// vba212ScanStats is deliberately caller-owned so tests and benchmarks can
// assert structural work without a process-global instrumentation hook.
type vba212ScanStats struct {
	RootTraversals int
	NodesVisited   int
}

// nonShortCircuitObjectGuardDocumentFindings walks the CST exactly once. Each
// candidate expression is attributed to its owning procedure through the
// byte ranges already recorded by procedure IR. This avoids the former
// procedures x document-nodes traversal while preserving outer-expression
// priority and per-procedure diagnostics.
func (a Analyzer) nonShortCircuitObjectGuardDocumentFindings(ctx context.Context, file parsedFile, procedures []sourceProcedure, stats *vba212ScanStats) ([]Finding, error) {
	projectEffects := effects.ProjectSummary{}
	if vba212SourceMayHaveGetter(file) {
		projectEffects = buildSingleFileProjectEffects(file)
	}
	return a.vba212ScanWithContext(ctx, file, procedures, stats, vba212Context{projectEffects: projectEffects})
}

func vba212ProcedureAtByte(procedures []sourceProcedure, startByte, endByte int) (sourceProcedure, bool) {
	i := sort.Search(len(procedures), func(i int) bool {
		return procedures[i].StartByte > startByte
	}) - 1
	if i < 0 {
		return sourceProcedure{}, false
	}
	proc := procedures[i]
	if startByte < proc.StartByte || endByte > proc.EndByte {
		return sourceProcedure{}, false
	}
	return proc, true
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

func physicalSourceLineCount(lines []string) int {
	count := len(lines)
	if count > 0 && lines[count-1] == "" {
		count--
	}
	return count
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

func resolveWithInfo(expr string, decls declarationScope) withInfo {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return withInfo{}
	}
	info := withInfo{Expression: expr}
	base := expr
	if idx := strings.Index(base, "("); idx >= 0 {
		base = base[:idx]
	}
	base = lastName(strings.TrimSpace(strings.TrimPrefix(base, "Set ")))
	if decl, ok := decls.lookup(base); ok {
		info.Target = base
		info.Type = decl.Type
	}
	return info
}

func currentWithInfo(stack []withInfo) (withInfo, bool) {
	for i := len(stack) - 1; i >= 0; i-- {
		if stack[i].Type != "" {
			return stack[i], true
		}
	}
	return withInfo{}, false
}

func currentWithExpression(stack []withInfo) string {
	for i := len(stack) - 1; i >= 0; i-- {
		if expression := strings.TrimSpace(stack[i].Expression); expression != "" {
			return expression
		}
	}
	return ""
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

func rangeFindAssignment(stmt, withExpression string) (string, string, bool) {
	left, _, ok := strings.Cut(stmt, "=")
	if !ok {
		return "", "", false
	}
	left = strings.TrimSpace(strings.TrimPrefix(strings.TrimSpace(left), "Set "))
	fields := strings.FieldsFunc(left, func(r rune) bool {
		return !isVBAIdentifierRune(r)
	})
	if len(fields) == 0 {
		return "", "", false
	}
	right := strings.TrimSpace(stmt[strings.Index(stmt, "=")+1:])
	memberIndex, ok := vbaMemberCallIndex(right, "Find")
	if !ok {
		return "", "", false
	}
	receiver := strings.TrimSpace(right[:memberIndex])
	if receiver == "" {
		receiver = strings.TrimSpace(withExpression)
		if receiver == "" {
			return "", "", false
		}
	}
	return cleanIdentifier(fields[len(fields)-1]), receiver, true
}

func (a Analyzer) rangeFindReceiverIsExcelRange(file parsedFile, receiver string, line int) bool {
	typ, ok := resolveExcelExpressionType(file, a.typeDB, receiver, line-1, a.RootDir, a.Config)
	return ok && isExcelRangeType(typ)
}

func vbaMemberCallIndex(text, member string) (int, bool) {
	lower := strings.ToLower(text)
	wanted := "." + strings.ToLower(member)
	for offset := 0; offset+len(wanted) <= len(lower); {
		index := strings.Index(lower[offset:], wanted)
		if index < 0 {
			return 0, false
		}
		index += offset
		if !vbaTextOffsetInString(text, index) {
			end := index + len(wanted)
			if end == len(lower) || lower[end] == '(' || unicode.IsSpace(rune(lower[end])) {
				return index, true
			}
		}
		offset = index + len(wanted)
	}
	return 0, false
}

func vbaTextOffsetInString(text string, offset int) bool {
	inString := false
	for i := 0; i < offset && i < len(text); i++ {
		if text[i] != '"' {
			continue
		}
		if inString && i+1 < len(text) && text[i+1] == '"' {
			i++
			continue
		}
		inString = !inString
	}
	return inString
}

func isCleanupFallthroughLabel(label string) bool {
	normalized := strings.ToLower(label)
	switch normalized {
	case "cleanup", "clean_up", "finally", "done":
		return true
	default:
		return strings.HasSuffix(normalized, "_cleanup") || strings.HasSuffix(normalized, "_clean_up")
	}
}

// isSemanticCleanupFallthrough recognizes handler labels that are intentionally
// shared by normal completion and error transfer. The name-based exception
// above remains the compatibility escape hatch, while this path covers
// qualified or project-specific labels whose body is demonstrably limited to
// resource cleanup or loop finalization. Arbitrary handler code is rejected so
// a normal fallthrough into logging or error reporting remains VBA204.
func isSemanticCleanupFallthrough(proc sourceProcedure, labelBlock vbacfg.Block, graph vbacfg.CFGView, loopExitVariables map[string]bool) bool {

	queue := []vbacfg.BlockID{labelBlock.ID}
	seen := map[vbacfg.BlockID]bool{}
	reachedNormalExit := false
	foundCleanup := false
	for len(queue) > 0 {
		blockID := queue[0]
		queue = queue[1:]
		if seen[blockID] {
			continue
		}
		seen[blockID] = true
		if blockID == graph.NormalExit() {
			reachedNormalExit = true
			continue
		}
		if blockID == graph.ExceptionalExit() || blockID == graph.TerminationExit() || blockID == graph.UnknownExit() {
			return false
		}
		block, ok := graph.BlockByID(blockID)
		if !ok {
			return false
		}
		if block.Statement != nil {
			if !isSemanticCleanupStatement(proc, *block.Statement, loopExitVariables) {
				return false
			}
			if isExecutableSemanticCleanup(proc, *block.Statement, loopExitVariables) {
				foundCleanup = true
			}
		}
		invalidEdge := false
		graph.ForEachOutgoing(blockID, func(edge vbacfg.Edge) bool {
			if edge.Class != vbacfg.EdgeNormal {
				return true
			}
			if edge.Kind == vbacfg.EdgeUnknown || edge.Uncertain {
				invalidEdge = true
				return false
			}
			queue = append(queue, edge.To)
			return true
		})
		if invalidEdge {
			return false
		}
	}
	return reachedNormalExit && foundCleanup
}

func isSemanticCleanupStatement(proc sourceProcedure, statement procedureir.Statement, loopExitVariables map[string]bool) bool {
	if statement.Recovered {
		return false
	}
	switch statement.Kind {
	case procedureir.StatementDeclaration, procedureir.StatementLabel,
		procedureir.StatementOnError, procedureir.StatementExit, procedureir.StatementEnd:
		return true
	case procedureir.StatementAssignment, procedureir.StatementSet:
		return isCleanupAssignment(statement) || isCleanupCall(proc, statement) || isLoopFinalizationAssignment(proc, statement, loopExitVariables)
	case procedureir.StatementCall:
		return isCleanupCall(proc, statement)
	case procedureir.StatementUnknown:
		return isCleanupAssignment(statement) || isCleanupCall(proc, statement)
	default:
		return false
	}
}

func isExecutableSemanticCleanup(proc sourceProcedure, statement procedureir.Statement, loopExitVariables map[string]bool) bool {
	if statement.Kind == procedureir.StatementDeclaration || statement.Kind == procedureir.StatementLabel || statement.Kind == procedureir.StatementOnError || statement.Kind == procedureir.StatementExit || statement.Kind == procedureir.StatementEnd {
		return false
	}
	if statement.Kind == procedureir.StatementAssignment || statement.Kind == procedureir.StatementSet {
		return isCleanupAssignment(statement) || isCleanupCall(proc, statement) || isLoopFinalizationAssignment(proc, statement, loopExitVariables)
	}
	return isCleanupCall(proc, statement)
}

func isCleanupAssignment(statement procedureir.Statement) bool {
	if statement.Value == nil || statement.Value.Recovered {
		return false
	}
	return strings.EqualFold(strings.TrimSpace(statement.Value.Text), "Nothing")
}

func isCleanupCall(proc sourceProcedure, statement procedureir.Statement) bool {
	for call := range proc.Calls.All() {
		if call.StatementID != statement.ID {
			continue
		}
		member := strings.TrimSpace(call.Callee.Member)
		if member == "" {
			member = strings.TrimSpace(call.Callee.BaseName)
		}
		member = strings.ToLower(cleanIdentifier(member))
		switch member {
		case "close", "pclose", "quit", "free", "release", "destroy":
			return true
		}
		// Declare Function aliases such as utc_pclose retain the cleanup
		// operation in a namespaced suffix. Do not generalize this to arbitrary
		// prefixes: CloseAndPurgeAll and ReleaseUserData are ordinary calls.
		if strings.HasSuffix(member, "_pclose") {
			return true
		}
	}
	// File-number Close is a VBA statement rather than a call expression, so
	// ProcedureIR exposes its exact syntax kind instead of a CallSite.
	return strings.EqualFold(statement.SyntaxKind, "close_statement")
}

func isLoopFinalizationAssignment(proc sourceProcedure, statement procedureir.Statement, loopExitVariables map[string]bool) bool {
	if statement.Kind != procedureir.StatementAssignment || statement.Target == nil || statement.Value == nil {
		return false
	}
	if len(loopExitVariables) == 0 {
		return false
	}
	target := strings.ToLower(strings.TrimSpace(statement.Target.Text))
	procedureName := strings.ToLower(strings.TrimSpace(proc.Name))
	if target != procedureName {
		return false
	}
	value := strings.ToLower(strings.TrimSpace(statement.Value.Text))
	if value == "" || value == "true" || value == "false" || value == "nothing" || value == "empty" || value == "null" {
		return false
	}
	if !strings.ContainsAny(value, "+-*/&") {
		return false
	}
	for variable := range loopExitVariables {
		if containsVBAIdentifier(value, variable) {
			return true
		}
	}
	return false
}

func containsVBAIdentifier(text, name string) bool {
	text = strings.ToLower(text)
	name = strings.ToLower(name)
	for start := 0; start <= len(text)-len(name); start++ {
		if !strings.HasPrefix(text[start:], name) {
			continue
		}
		end := start + len(name)
		if (start == 0 || !isIdentifierPart(text[start-1])) && (end == len(text) || !isIdentifierPart(text[end])) {
			return true
		}
	}
	return false
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
		if a.Code != b.Code {
			return a.Code < b.Code
		}
		if a.Procedure != b.Procedure {
			return a.Procedure < b.Procedure
		}
		return findingCycleSortKey(a) < findingCycleSortKey(b)
	})
}

func findingCycleSortKey(finding Finding) string {
	if finding.CallCycle == nil {
		return ""
	}
	parts := make([]string, 0, len(finding.CallCycle.Path))
	for _, node := range finding.CallCycle.Path {
		parts = append(parts, strings.ToLower(node.QualifiedName)+"|"+strings.ToLower(node.Kind)+"|"+node.File+"|"+strconvItoa(node.Line))
	}
	return strings.Join(parts, "->")
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
