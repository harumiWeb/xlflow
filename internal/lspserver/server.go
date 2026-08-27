package lspserver

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log"
	"math"
	"net/url"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"time"
	"unicode/utf16"

	"github.com/sourcegraph/jsonrpc2"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/analyze/semanticquery"
	"github.com/harumiWeb/xlflow/internal/config"
	formsintel "github.com/harumiWeb/xlflow/internal/excel/forms/intel"
	"github.com/harumiWeb/xlflow/internal/lint"
	staticrules "github.com/harumiWeb/xlflow/internal/staticanalysis/rules"
	"github.com/harumiWeb/xlflow/internal/typedb"
	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/constexpr"
	"github.com/harumiWeb/xlflow/internal/vba/effects"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	"github.com/harumiWeb/xlflow/internal/vbadb"
	"github.com/harumiWeb/xlflow/internal/vbafmt"
)

const serverName = "xlflow-vba-lsp"

const diagnosticsDebounce = 300 * time.Millisecond
const diagnosticsOpenDelay = 750 * time.Millisecond
const diagnosticsFullIdleDelay = 2 * time.Second
const diagnosticsLargeFileLines = 10_000

type BuildInfo struct {
	Version string
	Commit  string
	Date    string
}

type Options struct {
	RootDir        string
	Config         config.Config
	Build          BuildInfo
	LogFile        string
	Stderr         io.Writer
	TypeDBDir      string
	PerformanceLog bool
}

type Server struct {
	opts                      Options
	db                        *vbadb.DB
	analyzer                  intel.Analyzer
	handler                   protocol.Handler
	docs                      *documents
	logger                    *log.Logger
	performance               *performanceRecorder
	analysis                  *workspaceAnalysisIndex
	semanticTokens            *semanticTokenCache
	semanticTokenGenerator    func(intel.Document, []intel.Document) ([]intel.SemanticToken, error)
	codeLensConfig            intel.CodeLensConfig
	diagnostics               func(context.Context, intel.Document) []intel.Diagnostic
	diagnosticsRequest        func(context.Context, intel.DiagnosticRequest) intel.DiagnosticResult
	defaultDiagnosticsRequest uintptr
	projectDiagnosticsRequest func(context.Context, intel.DiagnosticRequest, intel.ProjectAnalysisSnapshot) intel.DiagnosticResult
	diagnosticsDebounce       time.Duration
	diagnosticsOpenDelay      time.Duration
	diagnosticsFullIdleDelay  time.Duration
	diagnosticsAfterFunc      func(time.Duration, func()) diagnosticTimer
	beforeDiagnosticsPublish  func()
	// performanceHook is an internal deterministic benchmark hook. It is never
	// set by production construction and therefore has no runtime cost there.
	performanceHook func(stage, path string)

	diagMu                   sync.Mutex
	diagStates               map[string]*diagnosticState
	diagWorkers              sync.WaitGroup
	diagStopped              bool
	analysisPermits          chan struct{}
	analysisPermitMu         sync.Mutex
	analysisPermitChanged    chan struct{}
	interactivePermitWaiters atomic.Int64
	overlayBuilds            atomic.Uint64
	overlayPublications      atomic.Uint64

	docLifecycleMu            sync.Mutex
	docLifecycles             map[string]*sync.Mutex
	projectSummaryCache       revisionCache[effects.ProjectSummary]
	resolutionCache           revisionCache[projectResolutionResult]
	projectConstantsCache     revisionCache[projectConstantsResult]
	semanticQueries           *semanticquery.Store
	semanticInvalidationMu    sync.Mutex
	semanticInvalidationPaths map[string]struct{}
	resolutionTypeLibSymbols  []procedureir.ResolverSymbol
}

type projectResolutionResult struct {
	resolver           procedureir.Resolver
	resolved           map[string]procedureir.DocumentIR
	diagnosticResolved map[string]procedureir.DocumentIR
}

type projectConstantsResult struct {
	visible map[string]bool
	values  map[string]constexpr.Value
}

type diagnosticTimer interface {
	Stop() bool
}

type diagnosticState struct {
	mu                   sync.Mutex
	generation           uint64
	latest               intel.Document
	notify               *glsp.Context
	timer                diagnosticTimer
	fullTimer            diagnosticTimer
	running              bool
	ready                bool
	readyMode            intel.DiagnosticMode
	runningMode          intel.DiagnosticMode
	publishedMode        intel.DiagnosticMode
	hasPublished         bool
	initialFast          bool
	open                 bool
	cancel               context.CancelFunc
	buildOverlay         bool
	runningOverlay       bool
	dependentPending     bool
	publishedSignatures  map[string]procedureSignature
	baselineKnown        bool
	diagnosticCache      *intel.DiagnosticCache
	changes              intel.ProcedureChangeSet
	overlayGeneration    uint64
	dependencyGeneration uint64
	projectReadyPending  bool
}

func Check(opts Options) error {
	result, err := typedb.LoadForRuntime(opts.TypeDBDir)
	if err != nil {
		return err
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	for _, warning := range result.Warnings {
		_, _ = fmt.Fprintf(w, "type database warning: %s\n", warning)
	}
	return intel.Analyzer{RootDir: opts.RootDir, Config: opts.Config, DB: result.DB}.Check()
}

func RunStdio(opts Options) error {
	s, cleanup, err := New(opts)
	if err != nil {
		return err
	}
	defer cleanup()
	stream := jsonrpc2.NewBufferedStream(stdioReadWriteCloser{}, jsonrpc2.VSCodeObjectCodec{})
	conn := jsonrpc2.NewConn(context.Background(), stream, rpcHandler{handler: &s.handler})
	<-conn.DisconnectNotify()
	return conn.Close()
}

func New(opts Options) (*Server, func(), error) {
	typeDB, err := typedb.LoadForRuntime(opts.TypeDBDir)
	if err != nil {
		return nil, nil, err
	}
	logger, cleanup, err := newLogger(opts)
	if err != nil {
		return nil, nil, err
	}
	for _, warning := range typeDB.Warnings {
		logger.Printf("type database warning: %s", warning)
	}
	docs := newDocuments(opts.RootDir, opts.Config.Src.Forms, opts.Config.Src.Workbook)
	docs.cfg = opts.Config
	s := &Server{
		opts:                     opts,
		db:                       typeDB.DB,
		resolutionTypeLibSymbols: workspaceTypeLibResolverSymbols(typeDB.DB),
		analyzer: intel.Analyzer{
			RootDir:                    opts.RootDir,
			Config:                     opts.Config,
			DB:                         typeDB.DB,
			TypeDBResolutionIncomplete: !typeDB.Complete,
			RealtimeFindingsFunc: func(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) ([]intel.RealtimeFinding, error) {
				findings, err := analyze.SourceRealtimeFindingsParsedIRCFGWithTypeDBContext(ctx, rootDir, cfg, doc, ir, controlFlow, typeDB.DB)
				if err != nil {
					return nil, err
				}
				out := make([]intel.RealtimeFinding, 0, len(findings))
				for _, finding := range findings {
					out = append(out, intel.RealtimeFinding{Code: finding.Code, Severity: finding.Severity, Line: finding.Line, Column: finding.Column, EndLine: finding.EndLine, EndColumn: finding.EndColumn, Message: finding.Message})
				}
				return out, nil
			},
		},
		docs:                      docs,
		logger:                    logger,
		performance:               newPerformanceRecorder(opts.PerformanceLog, logger),
		semanticTokens:            newSemanticTokenCache(),
		codeLensConfig:            intel.DefaultCodeLensConfig(),
		diagStates:                make(map[string]*diagnosticState),
		docLifecycles:             make(map[string]*sync.Mutex),
		semanticQueries:           semanticquery.New(semanticquery.Options{}),
		semanticInvalidationPaths: make(map[string]struct{}),
		analysisPermits:           make(chan struct{}, max(1, runtime.GOMAXPROCS(0)/2)),
		analysisPermitChanged:     make(chan struct{}),
	}
	s.analysis = s.newWorkspaceAnalysisIndex()
	s.analyzer.DocumentSymbolsFunc = s.cachedDocumentSourceSymbols
	s.analyzer.WorkspaceSymbolsFunc = s.cachedWorkspaceSymbols
	s.analyzer.WorkspaceSymbolQueryFunc = s.cachedWorkspaceSymbolQuery
	s.analyzer.WorkspaceSymbolsSnapshotFunc = s.cachedWorkspaceSymbolsSnapshot
	s.semanticTokenGenerator = s.analyzer.SemanticTokens
	s.diagnostics = s.analyzer.DiagnosticsContext
	s.diagnosticsRequest = s.analyzer.DiagnosticsRequestContext
	s.defaultDiagnosticsRequest = reflect.ValueOf(s.diagnosticsRequest).Pointer()
	s.projectDiagnosticsRequest = func(ctx context.Context, request intel.DiagnosticRequest, project intel.ProjectAnalysisSnapshot) intel.DiagnosticResult {
		ctx = semanticquery.WithContext(ctx, semanticquery.Context{
			Store:    s.semanticQueries,
			Revision: fmt.Sprintf("lsp-%d", project.Revision),
			Metrics:  analysisstats.FromContext(ctx),
		})
		initializeCapabilityTelemetry(ctx)
		projectByPath := make(map[string]intel.ProjectAnalysisDocument, len(project.Documents))
		capabilityDocuments := make([]analyze.ProjectCapabilityDocument, 0, len(project.Documents))
		for _, projectDocument := range project.Documents {
			projectByPath[symbolFileKey(projectDocument.IR.Path)] = projectDocument
			capabilityDocuments = append(capabilityDocuments, analyze.ProjectCapabilityDocument{
				IR: projectDocument.IR, CFG: projectDocument.CFG, Source: projectDocument.Source,
			})
		}
		resolutionComplete := project.Complete && typeDB.Complete
		resolutionResolver, resolvedProjectIR, resolvedDiagnosticIR := s.projectResolution(ctx, project, resolutionComplete)
		for i := range capabilityDocuments {
			if resolved, ok := resolvedProjectIR[symbolFileKey(capabilityDocuments[i].IR.Path)]; ok {
				capabilityDocuments[i].IR = resolved
			}
		}
		capabilityRequirements := analyze.PlanProjectCapabilities(s.opts.Config.Analyze, capabilityDocuments)
		var projectEffects effects.ProjectSummary
		if capabilityRequirements.Effects {
			projectEffects = s.projectEffectSummaryWithResolution(ctx, project, resolvedProjectIR, resolutionComplete)
		}
		// Resolution is built before planning because the planner needs its
		// completeness and resolved IR. ProjectConstants is currently part of
		// the unconditional compile-equivalent baseline, but keep this boundary
		// explicit so optional plans can skip it if that contract changes.
		var projectConstants projectConstantsResult
		if capabilityRequirements.ProjectConstants {
			projectConstants = s.projectConstants(project, resolutionComplete, typeDB.DB)
		}
		analyzer := s.analyzer
		analyzer.VisibleConstants = projectConstants.visible
		analyzer.ConstantValues = projectConstants.values
		analyzer.RealtimeFindingsFunc = func(ctx context.Context, rootDir string, cfg config.Config, doc *vbaast.ParsedDocument, ir procedureir.DocumentIR, controlFlow vbacfg.Document) ([]intel.RealtimeFinding, error) {
			projectKey := symbolFileKey(ir.Path)
			if projectDocument, ok := projectByPath[projectKey]; ok {
				if resolved, resolvedOK := resolvedProjectIR[projectKey]; resolvedOK {
					ir = resolved
				} else {
					ir = projectDocument.IR
				}
				controlFlow = projectDocument.CFG
			}
			findings, err := analyze.SourceRealtimeFindingsParsedIRCFGWithTypeDBAndProjectConstantsContext(ctx, rootDir, cfg, doc, ir, controlFlow, typeDB.DB, projectEffects, projectConstants.values)
			if err != nil {
				return nil, err
			}
			out := make([]intel.RealtimeFinding, 0, len(findings))
			for _, finding := range findings {
				out = append(out, intel.RealtimeFinding{Code: finding.Code, Severity: finding.Severity, Line: finding.Line, Column: finding.Column, EndLine: finding.EndLine, EndColumn: finding.EndColumn, Message: finding.Message})
			}
			// Compile-equivalent diagnostics use the full-resolution clone cached
			// with this project revision. The realtime/effects path keeps the
			// procedure-only view cached above for parity with batch analysis.
			var resolvedForDiagnostics procedureir.DocumentIR
			if cached, ok := resolvedDiagnosticIR[projectKey]; ok {
				resolvedForDiagnostics = cached
			} else {
				resolvedForDiagnostics = procedureir.Resolve(ir, resolutionResolver)
			}
			for _, diagnostic := range procedureir.Diagnostics(resolvedForDiagnostics, project.Complete && typeDB.Complete) {
				out = append(out, intel.RealtimeFinding{Code: diagnostic.Code, Severity: "error",
					Line: diagnostic.Range.StartLine, Column: diagnostic.Range.StartColumn,
					EndLine: diagnostic.Range.EndLine, EndColumn: diagnostic.Range.EndColumn, Message: diagnostic.Message})
			}
			return out, nil
		}
		return analyzer.DiagnosticsRequestContext(ctx, request)
	}
	s.diagnosticsDebounce = diagnosticsDebounce
	s.diagnosticsOpenDelay = diagnosticsOpenDelay
	s.diagnosticsFullIdleDelay = diagnosticsFullIdleDelay
	s.diagnosticsAfterFunc = func(delay time.Duration, callback func()) diagnosticTimer {
		return time.AfterFunc(delay, callback)
	}
	s.handler = protocol.Handler{
		Initialize:                          s.initialize,
		Initialized:                         s.initialized,
		Shutdown:                            s.shutdown,
		Exit:                                s.exit,
		TextDocumentDidOpen:                 s.didOpen,
		TextDocumentDidChange:               s.didChange,
		TextDocumentDidClose:                s.didClose,
		TextDocumentDocumentSymbol:          s.documentSymbol,
		WorkspaceSymbol:                     s.workspaceSymbol,
		WorkspaceDidChangeWatchedFiles:      s.didChangeWatchedFiles,
		TextDocumentDefinition:              s.definition,
		TextDocumentReferences:              s.references,
		TextDocumentPrepareCallHierarchy:    s.prepareCallHierarchy,
		CallHierarchyIncomingCalls:          s.callHierarchyIncomingCalls,
		CallHierarchyOutgoingCalls:          s.callHierarchyOutgoingCalls,
		TextDocumentPrepareRename:           s.prepareRename,
		TextDocumentRename:                  s.rename,
		TextDocumentHover:                   s.hover,
		TextDocumentCompletion:              s.completion,
		TextDocumentCodeAction:              s.codeAction,
		TextDocumentSignatureHelp:           s.signatureHelp,
		TextDocumentFormatting:              s.formatting,
		TextDocumentSemanticTokensFull:      s.semanticTokensFull,
		TextDocumentSemanticTokensFullDelta: s.semanticTokensFullDelta,
		TextDocumentCodeLens:                s.codeLens,
	}
	return s, func() {
		s.stopDiagnostics()
		if s.analysis != nil {
			s.analysis.stop()
		}
		s.docs.closeAll()
		cleanup()
	}, nil
}

func projectVisibleConstants(project intel.ProjectAnalysisSnapshot, typeDB *vbadb.DB) map[string]bool {
	counts := make(map[string]int)
	add := func(name string) {
		name = projectConstantIdentifier(name)
		if name != "" {
			counts[name]++
		}
	}
	addQualified := func(qualifier, name string) {
		qualifier = projectConstantIdentifier(qualifier)
		name = projectConstantIdentifier(name)
		if qualifier != "" && name != "" {
			counts[qualifier+"."+name]++
		}
	}
	for _, document := range project.Documents {
		standardModule := strings.EqualFold(document.IR.ModuleKind, "standard")
		for _, declaration := range document.IR.Declarations {
			if !declaration.IsConst && !procedureir.IsConstKind(declaration.Kind) {
				continue
			}
			if !strings.EqualFold(declaration.Visibility, "public") && !strings.EqualFold(declaration.Visibility, "friend") {
				continue
			}
			if standardModule {
				add(declaration.Name)
			}
			addQualified(document.IR.ModuleName, declaration.Name)
			addQualified(declaration.Parent, declaration.Name)
		}
	}
	if typeDB != nil {
		constants := typeDB.AllConstantsList()
		countsByName := make(map[string]int)
		for _, constant := range constants {
			name := projectConstantIdentifier(constant.Name)
			if name != "" {
				countsByName[name]++
			}
		}
		for _, constant := range constants {
			name := projectConstantIdentifier(constant.Name)
			if name == "" {
				continue
			}
			if countsByName[name] == 1 {
				counts[name]++
			}
			if group := projectConstantIdentifier(constant.EnumGroup); group != "" {
				counts[group+"."+name]++
			}
			if library := projectConstantIdentifier(constant.Library); library != "" {
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

func projectConstantValues(project intel.ProjectAnalysisSnapshot, typeDB *vbadb.DB) map[string]constexpr.Value {
	documents := make([]lint.ConstantValueDocument, 0, len(project.Documents))
	for _, document := range project.Documents {
		documents = append(documents, lint.ConstantValueDocument{Source: document.Source, IR: &document.IR})
	}
	return lint.ProjectConstantValues(documents, typeDB)
}

func projectConstantIdentifier(text string) string {
	text = strings.TrimSpace(text)
	text = strings.Trim(text, "[]")
	text = strings.TrimRight(text, "$%&#@^!")
	return strings.ToLower(text)
}

func (s *Server) projectEffectSummary(project intel.ProjectAnalysisSnapshot) effects.ProjectSummary {
	return s.projectEffectSummaryWithResolution(context.Background(), project, nil, false)
}

func (s *Server) projectEffectSummaryWithResolution(ctx context.Context, project intel.ProjectAnalysisSnapshot, resolved map[string]procedureir.DocumentIR, complete bool) effects.ProjectSummary {
	return s.projectSummaryCache.getOrBuild(project.Revision, complete, func() effects.ProjectSummary {
		measurement := s.performance.start("project/preparation", performanceStageProjectEffects, "interactive", "")
		finishCapability := analysisstats.MeasureCapabilityBuild(ctx, analysisstats.CapabilityEffectsBuildsCounter)
		documents := make([]effects.Document, len(project.Documents))
		for i, document := range project.Documents {
			ir := document.IR
			if resolvedIR, ok := resolved[symbolFileKey(document.IR.Path)]; ok {
				ir = resolvedIR
			}
			documents[i] = effects.Document{IR: ir, CFG: document.CFG}
		}
		summary := effects.Build(documents)
		finishCapability(nil)
		measurement.finish(summary.ProcedureCount(), 0, nil)
		return summary
	})
}

func initializeCapabilityTelemetry(ctx context.Context) {
	if recorder := analysisstats.FromContext(ctx); recorder != nil {
		for _, name := range analysisstats.CapabilityBuildCounters {
			// TypeDB is loaded once during Server construction, outside the
			// revision-scoped diagnostics recorder. Do not report a misleading
			// per-request zero for that server-lifetime capability.
			if name == analysisstats.CapabilityTypeDBBuildsCounter {
				continue
			}
			recorder.AddSum(name, 0)
		}
	}
}

func (s *Server) projectConstants(project intel.ProjectAnalysisSnapshot, complete bool, typeDB *vbadb.DB) projectConstantsResult {
	return s.projectConstantsCache.getOrBuild(project.Revision, complete, func() projectConstantsResult {
		measurement := s.performance.start("project/preparation", performanceStageProjectConstants, "interactive", "")
		result := projectConstantsResult{
			visible: projectVisibleConstants(project, typeDB),
			values:  projectConstantValues(project, typeDB),
		}
		measurement.finish(len(result.visible)+len(result.values), 0, nil)
		return result
	})
}

func newLogger(opts Options) (*log.Logger, func(), error) {
	if strings.TrimSpace(opts.LogFile) != "" {
		path := opts.LogFile
		if !filepath.IsAbs(path) {
			path = filepath.Join(opts.RootDir, path)
		}
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return nil, nil, err
		}
		file, err := os.OpenFile(path, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0o644)
		if err != nil {
			return nil, nil, err
		}
		return log.New(file, "xlflow-lsp: ", log.LstdFlags), func() { _ = file.Close() }, nil
	}
	w := opts.Stderr
	if w == nil {
		w = os.Stderr
	}
	return log.New(w, "xlflow-lsp: ", log.LstdFlags), func() {}, nil
}

func (s *Server) initialize(_ *glsp.Context, params *protocol.InitializeParams) (any, error) {
	s.codeLensConfig = codeLensConfigFromInitialize(params)
	capabilities := s.handler.CreateServerCapabilities()
	if capabilities.CodeLensProvider != nil {
		resolveProvider := false
		capabilities.CodeLensProvider.ResolveProvider = &resolveProvider
	}
	if syncOptions, ok := capabilities.TextDocumentSync.(*protocol.TextDocumentSyncOptions); ok {
		kind := protocol.TextDocumentSyncKindIncremental
		syncOptions.Change = &kind
	}
	if capabilities.CompletionProvider != nil {
		capabilities.CompletionProvider.TriggerCharacters = completionTriggerCharacters()
	}
	if capabilities.SignatureHelpProvider != nil {
		capabilities.SignatureHelpProvider.TriggerCharacters = []string{"(", ",", " "}
		capabilities.SignatureHelpProvider.RetriggerCharacters = []string{","}
	}
	if capabilities.RenameProvider != nil {
		prepareProvider := true
		capabilities.RenameProvider = protocol.RenameOptions{PrepareProvider: &prepareProvider}
	}
	if semantic, ok := capabilities.SemanticTokensProvider.(*protocol.SemanticTokensOptions); ok {
		semantic.Legend = protocol.SemanticTokensLegend{
			TokenTypes:     intel.SemanticTokenTypes,
			TokenModifiers: intel.SemanticTokenModifiers,
		}
		delta := true
		semantic.Full = protocol.SemanticDelta{Delta: &delta}
		semantic.Range = nil
	}
	version := s.opts.Build.Version
	if version == "" {
		version = "dev"
	}
	return protocol.InitializeResult{
		Capabilities: capabilities,
		ServerInfo: &protocol.InitializeResultServerInfo{
			Name:    serverName,
			Version: &version,
		},
	}, nil
}

func (s *Server) initialized(_ *glsp.Context, _ *protocol.InitializedParams) error {
	s.analysis.start()
	s.logger.Printf("initialized")
	return nil
}

func codeLensConfigFromInitialize(params *protocol.InitializeParams) intel.CodeLensConfig {
	cfg := intel.DefaultCodeLensConfig()
	if params == nil {
		return cfg
	}
	options, ok := params.InitializationOptions.(map[string]any)
	if !ok {
		return cfg
	}
	codeLens, ok := options["codeLens"].(map[string]any)
	if !ok {
		return cfg
	}
	if value, ok := codeLens["enabled"].(bool); ok {
		cfg.Enabled = value
	}
	if value, ok := codeLens["runProcedure"].(bool); ok {
		cfg.RunProcedure = value
	}
	if value, ok := codeLens["runTests"].(bool); ok {
		cfg.RunTests = value
	}
	if value, ok := codeLens["userFormEvents"].(bool); ok {
		cfg.UserFormEvents = value
	}
	return cfg
}

func (s *Server) shutdown(_ *glsp.Context) error {
	s.stopDiagnostics()
	if s.analysis != nil {
		s.analysis.stop()
	}
	s.docs.closeAll()
	s.logger.Printf("shutdown")
	return nil
}

func (s *Server) exit(_ *glsp.Context) error {
	s.logger.Printf("exit")
	return nil
}

func (s *Server) didOpen(ctx *glsp.Context, params *protocol.DidOpenTextDocumentParams) error {
	measurement := s.startPerformanceURI("textDocument/didOpen", string(params.TextDocument.URI))
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	doc, err := s.docs.open(uri, params.TextDocument.Text, int32(params.TextDocument.Version))
	if err != nil {
		unlock()
		measurement.finish(0, err)
		return err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) == DocumentKindVBA {
		s.semanticTokens.open(doc)
		s.semanticTokens.invalidateWorkspace()
	}
	s.openDiagnostics(ctx, doc)
	unlock()
	measurement.finish(0, nil)
	return nil
}

func (s *Server) didChange(ctx *glsp.Context, params *protocol.DidChangeTextDocumentParams) error {
	if len(params.ContentChanges) == 0 {
		return nil
	}
	changes, ok := decodeDocumentContentChanges(params.ContentChanges)
	if !ok {
		s.logger.Printf("ignored textDocument/didChange with unsupported content changes")
		return nil
	}
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	defer unlock()
	changeStarted := time.Now()
	change, err := s.docs.applyChangesWithResult(uri, changes, int32(params.TextDocument.Version))
	s.logDocumentChangePerformance(uri, int32(params.TextDocument.Version), change, changeStarted)
	if err != nil {
		return err
	}
	if !change.applied {
		s.logger.Printf("ignored textDocument/didChange for %q version=%d", uri, params.TextDocument.Version)
		return nil
	}
	if s.documentKind(change.document) == DocumentKindVBA {
		s.semanticTokens.invalidateWorkspace()
	}
	s.scheduleDiagnostics(ctx, change.document, diagnosticChangeSet(changes))
	return nil
}

type procedureSignature struct {
	name        string
	fingerprint string
}

func procedureSignaturesFromSymbols(syms []intel.Symbol) map[string]procedureSignature {
	out := make(map[string]procedureSignature)
	for _, sym := range syms {
		if !procedureSymbolKind(sym.Kind) {
			continue
		}
		var fingerprint strings.Builder
		fingerprint.WriteString(strings.ToLower(sym.Kind))
		fingerprint.WriteString("|")
		fingerprint.WriteString(strings.ToLower(sym.ReturnType))
		fingerprint.WriteString("|")
		fingerprint.WriteString(strings.ToLower(sym.Visibility))
		for _, param := range sym.Parameters {
			fmt.Fprintf(&fingerprint, "|%s:%s:%t:%s:%t:%t", strings.ToLower(param.Name), strings.ToLower(param.Type), param.IsArray, strings.ToLower(param.Passing), param.Optional, param.ParamArray)
		}
		key := strings.ToLower(sym.Module + "." + sym.Name + "." + sym.Kind)
		out[key] = procedureSignature{name: sym.Name, fingerprint: fingerprint.String()}
	}
	return out
}

func procedureSymbolKind(kind string) bool {
	switch strings.ToLower(kind) {
	case "sub", "function", "property", "property_get", "property_let", "property_set", "declare", "declare_sub", "declare_function":
		return true
	default:
		return false
	}
}

func changedProcedureNames(before, after map[string]procedureSignature) []string {
	seen := make(map[string]bool)
	var out []string
	for key, old := range before {
		current, exists := after[key]
		if exists && current.fingerprint == old.fingerprint {
			continue
		}
		if !seen[strings.ToLower(old.name)] {
			seen[strings.ToLower(old.name)] = true
			out = append(out, old.name)
		}
	}
	for key, current := range after {
		old, exists := before[key]
		if exists && old.fingerprint == current.fingerprint {
			continue
		}
		if !seen[strings.ToLower(current.name)] {
			seen[strings.ToLower(current.name)] = true
			out = append(out, current.name)
		}
	}
	return out
}

// scheduleByRefDependentDiagnostics refreshes open callers after a project
// procedure signature changes. The workspace index narrows work to calls with
// the changed base name, so unrelated open documents are not republished.
func (s *Server) scheduleByRefDependentDiagnostics(ctx *glsp.Context, changedURI string, names []string) {
	if !s.opts.Config.Analyze.DetectByRefArgumentMismatch || len(names) == 0 {
		return
	}
	openByPath := make(map[string]intel.Document)
	for _, doc := range s.docs.openDocuments() {
		if doc.URI != changedURI {
			openByPath[symbolFileKey(doc.Path)] = doc
		}
	}
	for _, name := range names {
		calls, err := s.analysis.queryResolvedCalls(workspaceCallQuery{CalleeBase: name})
		if err != nil {
			s.logger.Printf("ByRef caller refresh lookup failed for %q: %v", name, err)
			continue
		}
		for _, call := range calls {
			callPath := call.File
			if callPath != "" && !filepath.IsAbs(callPath) {
				callPath = filepath.Join(s.opts.RootDir, filepath.FromSlash(callPath))
			}
			if caller, ok := openByPath[symbolFileKey(callPath)]; ok {
				s.scheduleDiagnosticsOnly(ctx, caller)
			}
		}
	}
}

func (s *Server) scheduleProjectDependentDiagnostics(ctx *glsp.Context, changedURI string, paths []string) {
	if s.semanticQueries != nil && len(paths) > 0 {
		s.semanticInvalidationMu.Lock()
		if s.semanticInvalidationPaths == nil {
			s.semanticInvalidationPaths = make(map[string]struct{})
		}
		for _, path := range paths {
			s.semanticInvalidationPaths[path] = struct{}{}
		}
		s.semanticInvalidationMu.Unlock()
	}
	if !s.opts.Config.Analyze.DetectErrorSuppressionPropagation || len(paths) == 0 {
		return
	}
	openByPath := make(map[string]intel.Document)
	for _, doc := range s.docs.openDocuments() {
		if doc.URI != changedURI {
			openByPath[symbolFileKey(doc.Path)] = doc
		}
	}
	for _, path := range paths {
		if caller, ok := openByPath[symbolFileKey(path)]; ok {
			s.scheduleDiagnosticsOnly(ctx, caller)
		}
	}
}

func (s *Server) flushSemanticInvalidations(ctx context.Context) {
	if s.semanticQueries == nil {
		return
	}
	s.semanticInvalidationMu.Lock()
	paths := make([]string, 0, len(s.semanticInvalidationPaths))
	for path := range s.semanticInvalidationPaths {
		paths = append(paths, path)
	}
	s.semanticInvalidationPaths = make(map[string]struct{})
	s.semanticInvalidationMu.Unlock()
	if len(paths) > 0 {
		normalized := make([]string, 0, len(paths))
		seen := make(map[string]struct{}, len(paths))
		for _, path := range paths {
			path = symbolFileKey(path)
			if path == "" {
				continue
			}
			if _, ok := seen[path]; ok {
				continue
			}
			seen[path] = struct{}{}
			normalized = append(normalized, path)
		}
		if len(normalized) > 0 {
			s.semanticQueries.InvalidateProceduresContext(ctx, normalized...)
		}
	}
}

func (s *Server) scheduleProjectReadyDiagnosticsForCompleteProject(project intel.ProjectAnalysisSnapshot) {
	if !project.Complete {
		return
	}
	type pendingDiagnostic struct {
		notify *glsp.Context
		doc    intel.Document
	}
	pending := []pendingDiagnostic{}
	s.diagMu.Lock()
	for _, state := range s.diagStates {
		state.mu.Lock()
		if state.projectReadyPending && state.open {
			state.projectReadyPending = false
			pending = append(pending, pendingDiagnostic{notify: state.notify, doc: state.latest})
		}
		state.mu.Unlock()
	}
	s.diagMu.Unlock()
	for _, diagnostic := range pending {
		s.scheduleDiagnosticsOnly(diagnostic.notify, diagnostic.doc)
	}
}

func (s *Server) didClose(ctx *glsp.Context, params *protocol.DidCloseTextDocumentParams) error {
	uri := string(params.TextDocument.URI)
	unlock := s.lockDocumentLifecycle(uri)
	closingSignatures := s.closeDiagnostics(ctx, uri)
	if doc, err := s.docs.getOrRead(uri); err == nil && s.documentKind(doc) == DocumentKindVBA {
		s.semanticTokens.close(doc)
	}
	s.docs.close(uri)
	s.semanticTokens.invalidateWorkspace()
	path, pathErr := fileURIToPath(uri)
	refreshDisk := pathErr == nil && !isUserFormSpecPath(s.opts.RootDir, s.opts.Config.Src.Forms, path)
	unlock()
	if refreshDisk {
		s.scheduleCloseOverlayRefresh(ctx, uri, path, closingSignatures)
	}
	return nil
}

func (s *Server) scheduleCloseOverlayRefresh(ctx *glsp.Context, uri, path string, closingSignatures map[string]procedureSignature) {
	s.diagMu.Lock()
	if s.diagStopped {
		s.diagMu.Unlock()
		return
	}
	s.diagWorkers.Add(1)
	s.diagMu.Unlock()
	go func() {
		defer s.diagWorkers.Done()
		unlock := s.lockDocumentLifecycle(uri)
		if s.docs.isOpen(uri) {
			unlock()
			return
		}
		_, _ = s.analysis.projectChange()
		refresh := s.analysis.beginClearOverlay(path)
		unlock()
		disk, restored, err := s.analysis.finishClearOverlay(refresh)
		if err != nil {
			s.logger.Printf("workspace analysis index close refresh failed for %q: %v", path, err)
			return
		}
		var diskSignatures map[string]procedureSignature
		if restored {
			diskSignatures = procedureSignaturesFromSymbols(disk.symbols)
		}
		s.scheduleByRefDependentDiagnostics(ctx, uri, changedProcedureNames(closingSignatures, diskSignatures))
		project, impacted := s.analysis.projectChange()
		s.scheduleProjectDependentDiagnostics(ctx, uri, impacted)
		s.scheduleProjectReadyDiagnosticsForCompleteProject(project)
	}()
}

func (s *Server) didChangeWatchedFiles(ctx *glsp.Context, params *protocol.DidChangeWatchedFilesParams) error {
	_, _ = s.analysis.projectChange()
	for _, event := range params.Changes {
		path, err := fileURIToPath(string(event.URI))
		if err != nil {
			return err
		}
		if isUserFormSpecPath(s.opts.RootDir, s.opts.Config.Src.Forms, path) {
			s.docs.invalidateDisk(path)
			continue
		}
		paths, err := symbols.RelatedSourcePaths(s.opts.RootDir, s.opts.Config, path)
		if err != nil {
			return err
		}
		for _, affected := range paths {
			s.docs.invalidateDisk(affected)
			if err := s.analysis.updatePath(affected); err != nil {
				s.logger.Printf("workspace analysis index watcher update failed for %q: %v", affected, err)
				continue
			}
		}
	}
	project, impacted := s.analysis.projectChange()
	s.scheduleProjectDependentDiagnostics(ctx, "", impacted)
	s.scheduleProjectReadyDiagnosticsForCompleteProject(project)
	s.semanticTokens.invalidateWorkspace()
	return nil
}

func (s *Server) documentSymbol(_ *glsp.Context, params *protocol.DocumentSymbolParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/documentSymbol", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return []protocol.DocumentSymbol{}, nil
	}
	syms, err := s.analyzer.DocumentSymbols(doc)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.DocumentSymbol, 0, len(syms))
	for _, sym := range syms {
		detail := sym.Detail
		out = append(out, protocol.DocumentSymbol{
			Name:           sym.Name,
			Detail:         &detail,
			Kind:           symbolKind(sym.Kind),
			Range:          toProtocolRange(sym.Range),
			SelectionRange: toProtocolRange(sym.Selection),
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) workspaceSymbol(_ *glsp.Context, params *protocol.WorkspaceSymbolParams) ([]protocol.SymbolInformation, error) {
	measurement := s.startPerformance("workspace/symbol", intel.Document{})
	syms, err := s.analyzer.WorkspaceSymbols(s.docs.openDocuments(), params.Query)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.SymbolInformation, 0, len(syms))
	for _, sym := range syms {
		uri := sym.File
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(sym.File)
		}
		out = append(out, protocol.SymbolInformation{
			Name: sym.Name,
			Kind: symbolKind(sym.Kind),
			Location: protocol.Location{
				URI:   protocol.DocumentUri(uri),
				Range: toProtocolRange(sym.Selection),
			},
			ContainerName: &sym.Module,
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) definition(_ *glsp.Context, params *protocol.DefinitionParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/definition", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return []protocol.Location{}, nil
	}
	locs, err := s.analyzer.Definition(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), s.docs.uriForDisplayPath)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.Location, 0, len(locs))
	for _, loc := range locs {
		uri := loc.URI
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(loc.Path)
		}
		out = append(out, protocol.Location{URI: protocol.DocumentUri(uri), Range: toProtocolRange(loc.Range)})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) references(_ *glsp.Context, params *protocol.ReferenceParams) ([]protocol.Location, error) {
	measurement := s.startPerformanceURI("textDocument/references", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return []protocol.Location{}, nil
	}
	locs, err := s.analyzer.References(doc, fromProtocolPosition(params.Position), s.docs.openDocuments(), params.Context.IncludeDeclaration, s.docs.uriForDisplayPath)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.Location, 0, len(locs))
	for _, loc := range locs {
		uri := loc.URI
		if !strings.HasPrefix(uri, "file:") {
			uri = s.docs.uriForDisplayPath(loc.Path)
		}
		out = append(out, protocol.Location{URI: protocol.DocumentUri(uri), Range: toProtocolRange(loc.Range)})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) prepareRename(_ *glsp.Context, params *protocol.PrepareRenameParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if s.documentKind(doc) != DocumentKindVBA {
		return nil, nil
	}
	target, err := s.analyzer.PrepareRename(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil {
		return nil, err
	}
	return protocol.RangeWithPlaceholder{
		Range:       toProtocolRange(target.Range),
		Placeholder: target.Name,
	}, nil
}

func (s *Server) rename(_ *glsp.Context, params *protocol.RenameParams) (*protocol.WorkspaceEdit, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if s.documentKind(doc) != DocumentKindVBA {
		return &protocol.WorkspaceEdit{Changes: map[protocol.DocumentUri][]protocol.TextEdit{}}, nil
	}
	edits, err := s.analyzer.Rename(doc, fromProtocolPosition(params.Position), params.NewName, s.docs.openDocuments(), s.docs.uriForDisplayPath)
	if err != nil {
		return nil, err
	}
	changes := map[protocol.DocumentUri][]protocol.TextEdit{}
	for _, edit := range edits {
		uri := edit.URI
		if uri == "" {
			uri = s.docs.uriForDisplayPath(edit.Path)
		}
		docURI := protocol.DocumentUri(uri)
		changes[docURI] = append(changes[docURI], protocol.TextEdit{
			Range:   toProtocolRange(edit.Range),
			NewText: edit.NewText,
		})
	}
	return &protocol.WorkspaceEdit{Changes: changes}, nil
}

func (s *Server) hover(_ *glsp.Context, params *protocol.HoverParams) (*protocol.Hover, error) {
	measurement := s.startPerformanceURI("textDocument/hover", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) == DocumentKindUserFormYAML {
		hover := formsintel.HoverYAML(doc.Source, formsintel.Position{Line: int(params.Position.Line), Character: int(params.Position.Character)})
		if hover == nil {
			measurement.finish(0, nil)
			return nil, nil
		}
		measurement.finish(1, nil)
		return &protocol.Hover{
			Contents: protocol.MarkupContent{Kind: protocol.MarkupKindMarkdown, Value: hover.Contents},
			Range:    toProtocolRangePtr(intel.Range{Start: intel.Position{Line: hover.Range.Start.Line, Character: hover.Range.Start.Character}, End: intel.Position{Line: hover.Range.End.Line, Character: hover.Range.End.Character}}),
		}, nil
	}
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return nil, nil
	}
	hover, err := s.analyzer.Hover(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil || hover == nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.finish(1, nil)
	return &protocol.Hover{
		Contents: protocol.MarkupContent{
			Kind:  protocol.MarkupKindMarkdown,
			Value: hover.Contents,
		},
		Range: toProtocolRangePtr(hover.Range),
	}, nil
}

func (s *Server) completion(_ *glsp.Context, params *protocol.CompletionParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/completion", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) == DocumentKindUserFormYAML {
		items := userFormCompletionProtocolItems(formsintel.CompleteYAML(doc.Source, formsintel.Position{
			Line:      int(params.Position.Line),
			Character: int(params.Position.Character),
		}))
		measurement.finish(len(items), nil)
		return protocol.CompletionList{IsIncomplete: false, Items: items}, nil
	}
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return protocol.CompletionList{IsIncomplete: false, Items: []protocol.CompletionItem{}}, nil
	}
	completions, err := s.analyzer.Completions(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	items := make([]protocol.CompletionItem, 0, len(completions))
	for _, completion := range completions {
		kind := completionItemKind(completion.Kind)
		item := protocol.CompletionItem{
			Label: completion.Label,
			Kind:  &kind,
		}
		if completion.InsertText != "" {
			item.InsertText = &completion.InsertText
		}
		if completion.SortText != "" {
			item.SortText = &completion.SortText
		}
		if completion.ReplaceRange != nil {
			item.TextEdit = protocol.TextEdit{
				Range:   toProtocolRange(*completion.ReplaceRange),
				NewText: firstNonEmpty(completion.InsertText, completion.Label),
			}
		}
		if completion.Snippet {
			format := protocol.InsertTextFormatSnippet
			item.InsertTextFormat = &format
		}
		if completion.Detail != "" {
			item.Detail = &completion.Detail
		}
		if completion.Documentation != "" {
			item.Documentation = protocol.MarkupContent{
				Kind:  protocol.MarkupKindMarkdown,
				Value: completion.Documentation,
			}
		}
		items = append(items, item)
	}
	measurement.finish(len(items), nil)
	return protocol.CompletionList{IsIncomplete: false, Items: items}, nil
}

func userFormCompletionProtocolItems(completions []formsintel.CompletionItem) []protocol.CompletionItem {
	items := make([]protocol.CompletionItem, 0, len(completions))
	for _, completion := range completions {
		kind := userFormCompletionItemKind(completion.Kind)
		item := protocol.CompletionItem{
			Label: completion.Label,
			Kind:  &kind,
			TextEdit: protocol.TextEdit{
				Range:   toProtocolRange(userFormCompletionRange(completion.Replace)),
				NewText: completion.InsertText,
			},
		}
		if completion.SortText != "" {
			item.SortText = &completion.SortText
		}
		if completion.Snippet {
			format := protocol.InsertTextFormatSnippet
			item.InsertTextFormat = &format
		}
		if completion.Detail != "" {
			item.Detail = &completion.Detail
		}
		if completion.Documentation != "" {
			item.Documentation = protocol.MarkupContent{
				Kind:  protocol.MarkupKindMarkdown,
				Value: completion.Documentation,
			}
		}
		items = append(items, item)
	}
	return items
}

func userFormCompletionRange(r formsintel.Range) intel.Range {
	return intel.Range{
		Start: intel.Position{Line: r.Start.Line, Character: r.Start.Character},
		End:   intel.Position{Line: r.End.Line, Character: r.End.Character},
	}
}

func userFormCompletionItemKind(kind formsintel.CompletionKind) protocol.CompletionItemKind {
	switch kind {
	case formsintel.CompletionKindProperty:
		return protocol.CompletionItemKindProperty
	case formsintel.CompletionKindValue:
		return protocol.CompletionItemKindValue
	case formsintel.CompletionKindReference:
		return protocol.CompletionItemKindReference
	case formsintel.CompletionKindSnippet:
		return protocol.CompletionItemKindSnippet
	default:
		return protocol.CompletionItemKindText
	}
}

func (s *Server) codeAction(_ *glsp.Context, params *protocol.CodeActionParams) (any, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if s.documentKind(doc) != DocumentKindVBA {
		return []protocol.CodeAction{}, nil
	}
	actions, err := s.analyzer.CodeActions(doc, fromProtocolRange(params.Range))
	if err != nil {
		return nil, err
	}
	out := make([]protocol.CodeAction, 0, len(actions))
	requestURI := protocol.DocumentUri(params.TextDocument.URI)
	for _, action := range actions {
		kind := protocol.CodeActionKindRefactorRewrite
		if action.Kind == "quickfix" {
			kind = protocol.CodeActionKindQuickFix
		}
		out = append(out, protocol.CodeAction{
			Title: action.Title,
			Kind:  &kind,
			Edit: &protocol.WorkspaceEdit{Changes: map[protocol.DocumentUri][]protocol.TextEdit{
				requestURI: {{
					Range:   toProtocolRange(action.Range),
					NewText: action.NewText,
				}},
			}},
		})
	}
	return out, nil
}

func (s *Server) signatureHelp(_ *glsp.Context, params *protocol.SignatureHelpParams) (*protocol.SignatureHelp, error) {
	measurement := s.startPerformanceURI("textDocument/signatureHelp", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return nil, nil
	}
	help, err := s.analyzer.SignatureHelp(doc, fromProtocolPosition(params.Position), s.docs.openDocuments())
	if err != nil || help == nil || len(help.Signatures) == 0 {
		measurement.finish(0, err)
		return nil, err
	}
	activeSignature := protocol.UInteger(max(0, help.ActiveSignature))
	activeParameter := protocol.UInteger(max(0, help.ActiveParameter))
	signatures := make([]protocol.SignatureInformation, 0, len(help.Signatures))
	for _, sig := range help.Signatures {
		info := protocol.SignatureInformation{Label: sig.Label}
		if sig.Documentation != "" {
			info.Documentation = protocol.MarkupContent{Kind: protocol.MarkupKindMarkdown, Value: sig.Documentation}
		}
		for _, param := range sig.Parameters {
			paramInfo := protocol.ParameterInformation{Label: parameterLabel(param)}
			if param.Documentation != "" {
				paramInfo.Documentation = protocol.MarkupContent{Kind: protocol.MarkupKindMarkdown, Value: param.Documentation}
			}
			info.Parameters = append(info.Parameters, paramInfo)
		}
		signatures = append(signatures, info)
	}
	measurement.finish(len(signatures), nil)
	return &protocol.SignatureHelp{
		Signatures:      signatures,
		ActiveSignature: &activeSignature,
		ActiveParameter: &activeParameter,
	}, nil
}

func (s *Server) formatting(_ *glsp.Context, params *protocol.DocumentFormattingParams) ([]protocol.TextEdit, error) {
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		return nil, err
	}
	if s.documentKind(doc) != DocumentKindVBA {
		return []protocol.TextEdit{}, nil
	}
	if !documentSupportsFormatting(doc) {
		return []protocol.TextEdit{}, nil
	}
	formatted, err := vbafmt.FormatTextWithOptions(doc.Source, documentIsClass(doc), vbafmt.FormatConfig{
		LineNumbers:           vbafmt.LineNumberModePreserve,
		OperatorSpacing:       s.opts.Config.Fmt.OperatorSpacing,
		OperatorSpacingSet:    true,
		DeclarationSpacing:    s.opts.Config.Fmt.DeclarationSpacing,
		DeclarationSpacingSet: true,
		KeywordCasing:         s.opts.Config.Fmt.KeywordCasing,
		KeywordCasingSet:      true,
		BuiltinCasing:         s.opts.Config.Fmt.BuiltinCasing,
		BuiltinCasingSet:      true,
	})
	if err != nil {
		if vbafmt.IsFormatParseError(err) {
			return []protocol.TextEdit{}, nil
		}
		return nil, err
	}
	if formatted == doc.Source {
		return []protocol.TextEdit{}, nil
	}
	return []protocol.TextEdit{{
		Range:   fullDocumentRange(doc.Source),
		NewText: formatted,
	}}, nil
}

func (s *Server) semanticTokensFull(_ *glsp.Context, params *protocol.SemanticTokensParams) (*protocol.SemanticTokens, error) {
	measurement := s.startPerformanceURI("textDocument/semanticTokens/full", string(params.TextDocument.URI))
	if doc, err := s.docs.getOrRead(string(params.TextDocument.URI)); err != nil {
		measurement.finish(0, err)
		return nil, err
	} else if s.documentKind(doc) != DocumentKindVBA {
		measurement.setDocument(doc)
		measurement.finish(0, nil)
		return &protocol.SemanticTokens{Data: []protocol.UInteger{}}, nil
	}
	cacheStarted := time.Now()
	result, doc, hit, err := s.semanticTokenResult(string(params.TextDocument.URI))
	measurement.setDocument(doc)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	s.logDocumentCachePerformance("semanticTokens/cache", cacheStatus(hit), doc, len(result.data)/5, cacheStarted, nil)
	response := semanticTokensResponse(result)
	measurement.finish(len(result.data)/5, nil)
	return response, nil
}

// semanticTokensFullDelta returns a full result when the client's base is not
// retained or when the valid one-edit delta would cost at least as much on the
// wire as the complete result.
func (s *Server) semanticTokensFullDelta(_ *glsp.Context, params *protocol.SemanticTokensDeltaParams) (any, error) {
	measurement := s.startPerformanceURI("textDocument/semanticTokens/full/delta", string(params.TextDocument.URI))
	if doc, err := s.docs.getOrRead(string(params.TextDocument.URI)); err != nil {
		measurement.finish(0, err)
		return nil, err
	} else if s.documentKind(doc) != DocumentKindVBA {
		measurement.setDocument(doc)
		measurement.finish(0, nil)
		return &protocol.SemanticTokens{Data: []protocol.UInteger{}}, nil
	}
	cacheStarted := time.Now()
	result, doc, hit, err := s.semanticTokenResult(string(params.TextDocument.URI))
	measurement.setDocument(doc)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	s.logDocumentCachePerformance("semanticTokens/cache", cacheStatus(hit), doc, len(result.data)/5, cacheStarted, nil)
	full := semanticTokensResponse(result)
	previous, known := s.semanticTokens.previous(doc, params.PreviousResultID)
	if !known {
		measurement.finish(len(result.data)/5, nil)
		return full, nil
	}
	delta := &protocol.SemanticTokensDelta{
		ResultId: &result.resultID,
		Edits:    semanticTokenDeltaEdits(previous.data, result.data),
	}
	if semanticTokenResponseSize(delta) >= semanticTokenResponseSize(full) {
		measurement.finish(len(result.data)/5, nil)
		return full, nil
	}
	measurement.finish(len(result.data)/5, nil)
	return delta, nil
}

func (s *Server) semanticTokenResult(uri string) (cachedSemanticTokens, intel.Document, bool, error) {
	for {
		generation := s.semanticTokens.begin()
		doc, err := s.docs.getOrRead(uri)
		if err != nil {
			return cachedSemanticTokens{}, intel.Document{}, false, err
		}
		result, hit, err := s.semanticTokens.get(doc, generation, func() ([]protocol.UInteger, error) {
			tokens, err := s.semanticTokenGenerator(doc, s.docs.openDocuments())
			if err != nil {
				return nil, err
			}
			return encodeSemanticTokens(tokens), nil
		})
		if errors.Is(err, errSemanticTokensSuperseded) {
			continue
		}
		return result, doc, hit, err
	}
}

func semanticTokensResponse(result cachedSemanticTokens) *protocol.SemanticTokens {
	response := &protocol.SemanticTokens{Data: cloneSemanticTokenData(result.data)}
	if result.resultID != "" {
		response.ResultID = &result.resultID
	}
	return response
}

func semanticTokenDeltaEdits(previous, current []protocol.UInteger) []protocol.SemanticTokensEdit {
	prefix := 0
	for prefix < len(previous) && prefix < len(current) && previous[prefix] == current[prefix] {
		prefix++
	}
	suffix := 0
	for suffix < len(previous)-prefix && suffix < len(current)-prefix &&
		previous[len(previous)-1-suffix] == current[len(current)-1-suffix] {
		suffix++
	}
	if prefix == len(previous) && prefix == len(current) {
		return []protocol.SemanticTokensEdit{}
	}
	return []protocol.SemanticTokensEdit{{
		Start:       protocol.UInteger(prefix),
		DeleteCount: protocol.UInteger(len(previous) - prefix - suffix),
		Data:        cloneSemanticTokenData(current[prefix : len(current)-suffix]),
	}}
}

func semanticTokenResponseSize(response any) int {
	encoded, err := json.Marshal(response)
	if err != nil {
		return math.MaxInt
	}
	return len(encoded)
}

func (s *Server) codeLens(_ *glsp.Context, params *protocol.CodeLensParams) ([]protocol.CodeLens, error) {
	measurement := s.startPerformanceURI("textDocument/codeLens", string(params.TextDocument.URI))
	doc, err := s.docs.getOrRead(string(params.TextDocument.URI))
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	measurement.setDocument(doc)
	if s.documentKind(doc) != DocumentKindVBA {
		measurement.finish(0, nil)
		return []protocol.CodeLens{}, nil
	}
	procedures, err := s.analyzer.RunnableProcedures(doc, s.codeLensConfig)
	if err != nil {
		measurement.finish(0, err)
		return nil, err
	}
	out := make([]protocol.CodeLens, 0, len(procedures))
	for _, procedure := range procedures {
		title := "$(play) Run"
		command := "xlflow.runProcedure"
		if procedure.Kind == "test" {
			title = "$(beaker) Run Test"
			command = "xlflow.runTestProcedure"
		}
		pos := protocol.Position{Line: protocol.UInteger(max(0, procedure.Line)), Character: protocol.UInteger(max(0, procedure.Character))}
		out = append(out, protocol.CodeLens{
			Range: protocol.Range{Start: pos, End: pos},
			Command: &protocol.Command{
				Title:   title,
				Command: command,
				Arguments: []any{map[string]any{
					"uri":           procedure.URI,
					"name":          procedure.Name,
					"moduleName":    procedure.ModuleName,
					"qualifiedName": procedure.QualifiedName,
					"kind":          procedure.Kind,
					"moduleKind":    procedure.ModuleKind,
					"line":          procedure.Line,
					"character":     procedure.Character,
				}},
			},
		})
	}
	measurement.finish(len(out), nil)
	return out, nil
}

func (s *Server) openDiagnostics(ctx *glsp.Context, doc intel.Document) <-chan struct{} {
	s.diagMu.Lock()
	if s.diagStopped {
		s.diagMu.Unlock()
		return nil
	}
	state := s.diagStates[doc.URI]
	if state == nil {
		state = &diagnosticState{}
		s.diagStates[doc.URI] = state
	}
	state.mu.Lock()
	if state.open && sameScheduledDocument(state.latest, doc) {
		state.notify = ctx
		state.mu.Unlock()
		s.diagMu.Unlock()
		return nil
	}
	state.generation++
	generation := state.generation
	documentKind := s.documentKind(doc)
	largeFile := sourceLineCount(doc.Source) >= diagnosticsLargeFileLines
	fastFirst := largeFile && documentKind == DocumentKindVBA
	openDelay := time.Duration(0)
	if largeFile && !fastFirst {
		openDelay = s.diagnosticsOpenDelay
	}
	state.latest = doc
	state.notify = ctx
	state.ready = openDelay <= 0
	state.readyMode = intel.DiagnosticModeFull
	if fastFirst {
		state.ready = true
		state.readyMode = intel.DiagnosticModeFast
	}
	state.hasPublished = false
	state.initialFast = fastFirst
	state.changes = intel.ProcedureChangeSet{}
	state.open = true
	state.buildOverlay = documentKind == DocumentKindVBA
	if state.buildOverlay {
		_, _ = s.analysis.projectChange()
		previous, exists := s.analysis.beginOverlay(doc, generation)
		if exists {
			state.publishedSignatures = procedureSignaturesFromSymbols(previous.symbols)
			state.baselineKnown = true
		} else {
			state.publishedSignatures = nil
			state.baselineKnown = false
		}
	}
	if state.timer != nil {
		state.timer.Stop()
		state.timer = nil
	}
	if state.fullTimer != nil {
		state.fullTimer.Stop()
		state.fullTimer = nil
	}
	if state.cancel != nil {
		state.cancel()
	}
	if fastFirst {
		state.fullTimer = s.diagnosticsAfterFunc(s.diagnosticsOpenDelay, func() {
			s.diagnosticsReady(doc.URI, state, generation, intel.DiagnosticModeFull)
		})
	}
	state.mu.Unlock()
	s.diagMu.Unlock()
	if openDelay > 0 {
		state.mu.Lock()
		if state.open && state.generation == generation && state.timer == nil {
			state.timer = s.diagnosticsAfterFunc(openDelay, func() {
				s.diagnosticsReady(doc.URI, state, generation, intel.DiagnosticModeFull)
			})
		}
		state.mu.Unlock()
		return nil
	}
	return s.launchDiagnostics(doc.URI, state)
}

func diagnosticChangeSet(changes []documentContentChange) intel.ProcedureChangeSet {
	result := intel.ProcedureChangeSet{}
	for _, change := range changes {
		if change.rng == nil {
			return intel.ProcedureChangeSet{Ranges: []intel.Range{{Start: intel.Position{}, End: intel.Position{Line: math.MaxInt}}}}
		}
		start := intel.Position{Line: int(change.rng.Start.Line), Character: int(change.rng.Start.Character)}
		end := intel.Position{Line: int(change.rng.End.Line), Character: int(change.rng.End.Character)}
		insertedLines := strings.Count(change.text, "\n")
		if insertedLines > 0 {
			end.Line = max(end.Line, start.Line+insertedLines)
		}
		result.Ranges = append(result.Ranges, intel.Range{Start: start, End: end})
	}
	return result
}

func (s *Server) scheduleDiagnostics(ctx *glsp.Context, doc intel.Document, changes intel.ProcedureChangeSet) {
	s.scheduleDocumentAnalysis(ctx, doc, true, changes)
}

func (s *Server) scheduleDiagnosticsOnly(ctx *glsp.Context, doc intel.Document) {
	s.scheduleDocumentAnalysis(ctx, doc, false, intel.ProcedureChangeSet{})
}

func (s *Server) scheduleDocumentAnalysis(ctx *glsp.Context, doc intel.Document, buildOverlay bool, changes intel.ProcedureChangeSet) {
	s.diagMu.Lock()
	if s.diagStopped {
		s.diagMu.Unlock()
		return
	}
	state := s.diagStates[doc.URI]
	if state == nil {
		state = &diagnosticState{open: true}
		s.diagStates[doc.URI] = state
	}
	state.mu.Lock()
	if buildOverlay && state.open && sameScheduledDocument(state.latest, doc) {
		state.notify = ctx
		state.mu.Unlock()
		s.diagMu.Unlock()
		return
	}
	// A dependent refresh must not supersede a source generation before its
	// overlay is published. Queue one follow-up diagnostic pass instead.
	if !buildOverlay && (state.buildOverlay || (state.running && state.runningOverlay)) {
		state.dependencyGeneration++
		state.dependentPending = true
		state.notify = ctx
		state.mu.Unlock()
		s.diagMu.Unlock()
		return
	}
	state.generation++
	if !buildOverlay {
		state.dependencyGeneration++
	}
	generation := state.generation
	fastFull := buildOverlay && s.documentKind(doc) == DocumentKindVBA
	state.latest = doc
	state.notify = ctx
	state.ready = false
	state.readyMode = intel.DiagnosticModeFull
	if fastFull {
		state.readyMode = intel.DiagnosticModeFast
	}
	state.publishedMode = intel.DiagnosticModeFull
	state.hasPublished = false
	state.initialFast = false
	state.changes = changes
	state.open = true
	state.buildOverlay = buildOverlay && s.documentKind(doc) == DocumentKindVBA
	if state.buildOverlay {
		_, _ = s.analysis.projectChange()
		if previous, exists := s.analysis.beginOverlay(doc, generation); exists {
			state.publishedSignatures = procedureSignaturesFromSymbols(previous.symbols)
			state.baselineKnown = true
		}
	}
	if state.timer != nil {
		state.timer.Stop()
	}
	if state.fullTimer != nil {
		state.fullTimer.Stop()
	}
	if state.cancel != nil {
		state.cancel()
	}
	state.timer = s.diagnosticsAfterFunc(s.diagnosticsDebounce, func() {
		mode := intel.DiagnosticModeFull
		if fastFull {
			mode = intel.DiagnosticModeFast
		}
		s.diagnosticsReady(doc.URI, state, generation, mode)
	})
	if fastFull {
		state.fullTimer = s.diagnosticsAfterFunc(s.diagnosticsFullIdleDelay, func() {
			s.diagnosticsReady(doc.URI, state, generation, intel.DiagnosticModeFull)
		})
	}
	state.mu.Unlock()
	s.diagMu.Unlock()
}

func (s *Server) diagnosticsReady(uri string, state *diagnosticState, generation uint64, mode intel.DiagnosticMode) {
	state.mu.Lock()
	if !state.open || state.generation != generation {
		state.mu.Unlock()
		return
	}
	if mode == intel.DiagnosticModeFast {
		state.timer = nil
	} else {
		state.fullTimer = nil
	}
	state.ready = true
	if mode == intel.DiagnosticModeFull || state.readyMode != intel.DiagnosticModeFull {
		state.readyMode = mode
	}
	state.mu.Unlock()
	s.launchDiagnostics(uri, state)
}

func (s *Server) launchDiagnostics(uri string, state *diagnosticState) <-chan struct{} {
	s.diagMu.Lock()
	if s.diagStopped || s.diagStates[uri] != state {
		s.diagMu.Unlock()
		return nil
	}
	state.mu.Lock()
	if !state.open || !state.ready || state.running {
		state.mu.Unlock()
		s.diagMu.Unlock()
		return nil
	}
	doc := state.latest
	notify := state.notify
	generation := state.generation
	buildOverlay := state.buildOverlay
	mode := state.readyMode
	runCtx, cancel := context.WithCancel(context.Background())
	state.ready = false
	state.running = true
	state.runningMode = mode
	state.buildOverlay = false
	state.runningOverlay = buildOverlay
	state.cancel = cancel
	s.diagWorkers.Add(1)
	done := make(chan struct{})
	state.mu.Unlock()
	s.diagMu.Unlock()

	go func() {
		defer close(done)
		s.runDocumentAnalysis(runCtx, uri, state, generation, doc, notify, buildOverlay, mode)
	}()
	return done
}

func sameScheduledDocument(left, right intel.Document) bool {
	if left.URI != right.URI || left.Path != right.Path || left.Version != right.Version || left.ModuleKind != right.ModuleKind {
		return false
	}
	if left.Snapshot != nil || right.Snapshot != nil {
		return left.Snapshot != nil && left.Snapshot == right.Snapshot &&
			left.Snapshot.Matches(left) && right.Snapshot.Matches(right)
	}
	return left.Source == right.Source
}

func (s *Server) runDocumentAnalysis(
	runCtx context.Context,
	uri string,
	state *diagnosticState,
	generation uint64,
	doc intel.Document,
	notify *glsp.Context,
	buildOverlay bool,
	mode intel.DiagnosticMode,
) {
	defer s.diagWorkers.Done()
	release, _, err := s.acquireAnalysisPermit(runCtx, "interactive")
	if err != nil {
		s.finishDocumentAnalysis(uri, state)
		return
	}
	defer release()

	if mode == intel.DiagnosticModeFast {
		s.runDiagnosticsBody(runCtx, uri, state, generation, doc, notify, mode)
	}

	if buildOverlay && runCtx.Err() == nil {
		state.mu.Lock()
		baselineKnown := state.baselineKnown
		state.mu.Unlock()
		if !baselineKnown {
			baseline, baselineErr := s.diskProcedureSignatures(runCtx, doc.Path)
			if baselineErr != nil && runCtx.Err() == nil {
				s.logger.Printf("workspace analysis disk signature baseline failed for %q: %v", doc.Path, baselineErr)
			}
			state.mu.Lock()
			if state.open && state.generation == generation && !state.baselineKnown {
				state.publishedSignatures = baseline
				state.baselineKnown = baselineErr == nil
			}
			state.mu.Unlock()
		}
		started := time.Now()
		s.overlayBuilds.Add(1)
		declarations, overlayDoc, included, err := s.analyzeWorkspaceOverlayDeclarations(runCtx, doc)
		if overlayDoc.Snapshot != nil && overlayDoc.Snapshot != doc.Snapshot {
			defer overlayDoc.Snapshot.Retire()
		}
		publishedDeclarations := false
		if err == nil && included && runCtx.Err() == nil && s.analysisGenerationCurrent(state, generation, doc) {
			publishedDeclarations = s.analysis.publishOverlayDeclarations(doc, generation, declarations)
		}
		if !publishedDeclarations && runCtx.Err() == nil && s.analysisGenerationCurrent(state, generation, doc) {
			s.analysis.abandonOverlay(doc, generation)
		}
		if !publishedDeclarations {
			s.logWorkspaceOverlayPerformance(doc, generation, started, err, true)
		}
		if publishedDeclarations && runCtx.Err() == nil {
			analysis, semanticIncluded, semanticErr := s.analyzeWorkspaceOverlay(runCtx, overlayDoc)
			publishedSemantic := false
			if semanticErr == nil && semanticIncluded && runCtx.Err() == nil && s.analysisGenerationCurrent(state, generation, doc) {
				publishedSemantic = s.analysis.publishOverlaySemantic(doc, generation, analysis)
			}
			s.logWorkspaceOverlayPerformance(doc, generation, started, semanticErr, !publishedSemantic)
			if !publishedSemantic && runCtx.Err() == nil && s.analysisGenerationCurrent(state, generation, doc) {
				s.analysis.abandonOverlay(doc, generation)
			}
			if publishedSemantic {
				s.overlayPublications.Add(1)
				newSignatures := procedureSignaturesFromSymbols(declarations.symbols)
				state.mu.Lock()
				oldSignatures := state.publishedSignatures
				if state.open && state.generation == generation {
					state.publishedSignatures = newSignatures
					state.overlayGeneration = generation
				}
				state.mu.Unlock()
				s.scheduleByRefDependentDiagnostics(notify, doc.URI, changedProcedureNames(oldSignatures, newSignatures))
				project, impacted := s.analysis.projectChangeClass("interactive")
				s.scheduleProjectDependentDiagnostics(notify, doc.URI, impacted)
				s.scheduleProjectReadyDiagnosticsForCompleteProject(project)
			}
		}
	}
	// Overlay failure deliberately does not suppress file-local diagnostics.
	if mode == intel.DiagnosticModeFull && runCtx.Err() == nil {
		s.runDiagnosticsBody(runCtx, uri, state, generation, doc, notify, mode)
	}
	s.completeDocumentAnalysis(uri, state, generation, mode)
}

func (s *Server) analysisGenerationCurrent(state *diagnosticState, generation uint64, doc intel.Document) bool {
	state.mu.Lock()
	defer state.mu.Unlock()
	return state.open && state.generation == generation && sameScheduledDocument(state.latest, doc)
}

func (s *Server) finishDocumentAnalysis(uri string, state *diagnosticState) {
	state.mu.Lock()
	state.running = false
	state.runningOverlay = false
	state.cancel = nil
	ready := state.open && state.ready
	state.mu.Unlock()
	if ready {
		s.launchDiagnostics(uri, state)
	}
}

func (s *Server) completeDocumentAnalysis(uri string, state *diagnosticState, generation uint64, mode intel.DiagnosticMode) {
	state.mu.Lock()
	state.running = false
	state.runningOverlay = false
	state.cancel = nil
	if mode == intel.DiagnosticModeFull && state.open && state.generation == generation && state.dependentPending {
		state.dependentPending = false
		state.generation++
		state.ready = true
		state.readyMode = intel.DiagnosticModeFull
	}
	ready := state.open && state.ready
	state.mu.Unlock()
	if ready {
		s.launchDiagnostics(uri, state)
	}
}

func (s *Server) runDiagnosticsBody(
	runCtx context.Context,
	uri string,
	state *diagnosticState,
	generation uint64,
	doc intel.Document,
	notify *glsp.Context,
	mode intel.DiagnosticMode,
) {
	measurement := s.startPerformance("diagnostics", doc)
	phase := "full"
	if mode == intel.DiagnosticModeFast {
		phase = "fast"
	}
	measurement.setPhase(phase)
	s.notifyPerformanceHook("diagnostics-"+phase+"-start", doc.Path)
	defer s.notifyPerformanceHook("diagnostics-"+phase+"-end", doc.Path)
	counter := performanceCounterFullDiagnosticRuns
	if mode == intel.DiagnosticModeFast {
		counter = performanceCounterFastDiagnosticRuns
	}
	s.performance.addCounter(counter, 1, "diagnostics", "diagnostics", "interactive", doc.Path)
	recorder := analysisstats.NewRecorder()
	runCtx = analysisstats.WithRecorder(runCtx, recorder)
	runCtx = semanticquery.WithContext(runCtx, semanticquery.Context{Store: s.semanticQueries, Metrics: recorder})
	s.flushSemanticInvalidations(runCtx)
	if mode == intel.DiagnosticModeFull {
		initializeCapabilityTelemetry(runCtx)
	}
	state.mu.Lock()
	previousCache := state.diagnosticCache
	changes := state.changes
	initialFast := state.initialFast
	dependencyGeneration := state.dependencyGeneration
	state.mu.Unlock()
	var result intel.DiagnosticResult
	if s.documentKind(doc) == DocumentKindVBA && s.diagnosticsRequest != nil {
		request := intel.DiagnosticRequest{Document: doc, Mode: mode, Changes: changes, PreviousCache: previousCache, Recorder: recorder, InitialFast: initialFast && mode == intel.DiagnosticModeFast}
		project := intel.ProjectAnalysisSnapshot{}
		if mode == intel.DiagnosticModeFull {
			project = s.analysis.projectSnapshotClass("interactive")
			if !project.Complete && s.projectDiagnosticsEnabled() {
				s.scheduleProjectReadyDiagnostics(state)
			}
		}
		if mode == intel.DiagnosticModeFull && project.Complete && s.projectDiagnosticsEnabled() {
			result = s.projectDiagnosticsRequest(runCtx, request, project)
		} else {
			result = s.diagnosticsRequest(runCtx, request)
		}
		if mode == intel.DiagnosticModeFast || (mode == intel.DiagnosticModeFull && !project.Complete) {
			// Project-negative resolution is fail-open while the workspace index
			// is pending, stale, or backed by recovered IR.  Keep the existing
			// interprocedural suppression and apply the same policy to VB052-
			// VB054 so Fast and incomplete Full diagnostics cannot leak guesses.
			result.Diagnostics = diagnosticsWithoutCodes(result.Diagnostics, "VBA237", "VB052", "VB053", "VB054")
		}
	} else {
		result.Diagnostics = s.documentDiagnostics(runCtx, doc)
	}
	diagnostics := result.Diagnostics
	out := make([]protocol.Diagnostic, 0, len(diagnostics))
	for _, diag := range diagnostics {
		out = append(out, toProtocolDiagnostic(diag))
	}
	if s.beforeDiagnosticsPublish != nil {
		s.beforeDiagnosticsPublish()
	}

	state.mu.Lock()
	discarded := !state.open || state.generation != generation || state.dependencyGeneration != dependencyGeneration || runCtx.Err() != nil || (mode == intel.DiagnosticModeFull && state.dependentPending) || (mode == intel.DiagnosticModeFast && state.hasPublished && state.publishedMode == intel.DiagnosticModeFull)
	if !discarded && notify != nil {
		notify.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
			URI:         protocol.DocumentUri(doc.URI),
			Diagnostics: out,
		})
	}
	if !discarded {
		state.publishedMode = mode
		state.hasPublished = true
		if mode == intel.DiagnosticModeFull && result.Cache != nil {
			state.diagnosticCache = result.Cache
		}
	}
	state.mu.Unlock()
	measurement.finishDiagnostics(len(out), generation, discarded)
	s.logDiagnosticStages(doc, generation, mode, recorder)
}

func (s *Server) projectDiagnosticsEnabled() bool {
	return s.projectDiagnosticsRequest != nil && s.diagnosticsRequest != nil &&
		reflect.ValueOf(s.diagnosticsRequest).Pointer() == s.defaultDiagnosticsRequest
}

func (s *Server) scheduleProjectReadyDiagnostics(state *diagnosticState) {
	state.mu.Lock()
	if state.projectReadyPending || !state.open {
		state.mu.Unlock()
		return
	}
	state.projectReadyPending = true
	state.mu.Unlock()
	if project := s.analysis.projectSnapshotClass("interactive"); project.Complete {
		s.scheduleProjectReadyDiagnosticsForCompleteProject(project)
		return
	}
	if s.analysis.initialReady() {
		return
	}
	s.diagWorkers.Add(1)
	go func() {
		defer s.diagWorkers.Done()
		err := s.analysis.waitReady()
		if err != nil {
			state.mu.Lock()
			state.projectReadyPending = false
			state.mu.Unlock()
			s.logger.Printf("workspace project diagnostics readiness failed: %v", err)
			return
		}
		project, _ := s.analysis.projectChangeClass("background")
		s.scheduleProjectReadyDiagnosticsForCompleteProject(project)
	}()
}

func diagnosticsWithoutCode(in []intel.Diagnostic, code string) []intel.Diagnostic {
	return diagnosticsWithoutCodes(in, code)
}

func diagnosticsWithoutCodes(in []intel.Diagnostic, codes ...string) []intel.Diagnostic {
	blocked := make(map[string]struct{}, len(codes))
	for _, code := range codes {
		blocked[strings.ToUpper(code)] = struct{}{}
	}
	out := make([]intel.Diagnostic, 0, len(in))
	for _, diagnostic := range in {
		if _, skip := blocked[strings.ToUpper(diagnostic.Code)]; !skip {
			out = append(out, diagnostic)
		}
	}
	return out
}

func toProtocolDiagnostic(diag intel.Diagnostic) protocol.Diagnostic {
	severityName := diag.Severity
	metadata, hasMetadata := staticrules.Lookup(diag.Code)
	if hasMetadata {
		severityName = string(metadata.DefaultSeverity)
		requestedName := strings.ToLower(diag.Severity)
		if requestedName == "info" {
			requestedName = "information"
		}
		requested := staticrules.RuleSeverity(requestedName)
		for _, supported := range metadata.SupportedSeverities {
			if requested == supported {
				severityName = string(requested)
				break
			}
		}
	}
	severity := diagnosticSeverity(severityName)
	source := diag.Source
	code := protocol.IntegerOrString{Value: diag.Code}
	out := protocol.Diagnostic{
		Range:    toProtocolRange(diag.Range),
		Severity: &severity,
		Code:     &code,
		Source:   &source,
		Message:  diag.Message,
	}
	if hasMetadata {
		out.CodeDescription = &protocol.CodeDescription{HRef: protocol.URI(metadata.DocumentationURL)}
	}
	return out
}

func (s *Server) documentKind(doc intel.Document) DocumentKind {
	return s.docs.documentKind(doc)
}

func (s *Server) documentDiagnostics(ctx context.Context, doc intel.Document) []intel.Diagnostic {
	switch s.documentKind(doc) {
	case DocumentKindVBA:
		return s.diagnostics(ctx, doc)
	case DocumentKindUserFormYAML:
		return userFormYAMLDiagnostics(doc)
	default:
		return nil
	}
}

var yamlErrorLocation = regexp.MustCompile(`line (\d+)(?:: column (\d+))?`)

func yamlErrorPosition(source string, err error) (int, int) {
	matches := yamlErrorLocation.FindStringSubmatch(err.Error())
	if len(matches) < 2 {
		return 0, 0
	}
	line, _ := strconv.Atoi(matches[1])
	line = max(0, line-1)
	column := 0
	if len(matches) > 2 && matches[2] != "" {
		column, _ = strconv.Atoi(matches[2])
		column = max(0, column-1)
	}
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	if line >= len(lines) {
		return max(0, len(lines)-1), 0
	}
	column = min(column, len(lines[line]))
	return line, utf16Len(lines[line][:column])
}

func (s *Server) closeDiagnostics(ctx *glsp.Context, uri string) map[string]procedureSignature {
	s.diagMu.Lock()
	state := s.diagStates[uri]
	s.diagMu.Unlock()
	if state != nil {
		state.mu.Lock()
		signatures := state.publishedSignatures
		state.close()
		if ctx != nil {
			ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
				URI:         protocol.DocumentUri(uri),
				Diagnostics: []protocol.Diagnostic{},
			})
		}
		state.mu.Unlock()
		return signatures
	} else if ctx != nil {
		ctx.Notify(string(protocol.ServerTextDocumentPublishDiagnostics), protocol.PublishDiagnosticsParams{
			URI:         protocol.DocumentUri(uri),
			Diagnostics: []protocol.Diagnostic{},
		})
	}
	return nil
}

func (state *diagnosticState) close() {
	state.generation++
	state.open = false
	state.ready = false
	state.latest = intel.Document{}
	state.notify = nil
	state.buildOverlay = false
	state.runningOverlay = false
	state.dependentPending = false
	state.publishedSignatures = nil
	state.baselineKnown = false
	state.diagnosticCache = nil
	state.hasPublished = false
	state.initialFast = false
	state.overlayGeneration = 0
	state.dependencyGeneration++
	state.projectReadyPending = false
	state.changes = intel.ProcedureChangeSet{}
	if state.timer != nil {
		state.timer.Stop()
		state.timer = nil
	}
	if state.fullTimer != nil {
		state.fullTimer.Stop()
		state.fullTimer = nil
	}
	if state.cancel != nil {
		state.cancel()
	}
}

func (s *Server) stopDiagnostics() {
	s.diagMu.Lock()
	var states []*diagnosticState
	if !s.diagStopped {
		s.diagStopped = true
		for uri, state := range s.diagStates {
			states = append(states, state)
			delete(s.diagStates, uri)
		}
	}
	s.diagMu.Unlock()
	for _, state := range states {
		state.mu.Lock()
		state.close()
		state.mu.Unlock()
	}
	s.diagWorkers.Wait()
}

func (s *Server) lockDocumentLifecycle(uri string) func() {
	s.docLifecycleMu.Lock()
	lifecycle := s.docLifecycles[uri]
	if lifecycle == nil {
		lifecycle = &sync.Mutex{}
		s.docLifecycles[uri] = lifecycle
	}
	s.docLifecycleMu.Unlock()
	lifecycle.Lock()
	return lifecycle.Unlock
}

func (s *Server) newWorkspaceAnalysisIndex() *workspaceAnalysisIndex {
	index := newWorkspaceAnalysisIndex(s.opts.RootDir, s.opts.Config, s.parseIndexedFileContext, s.logInitialWorkspaceIndexPerformance)
	index.nonBlockingQueries = true
	index.performance = s.performance
	index.logDeclarations = s.logInitialWorkspaceDeclarationPerformance
	index.parseDeclarations = s.parseIndexedDeclarationsContext
	if permits := cap(s.analysisPermits); permits > 0 {
		index.initialWorkers = permits
		index.semanticWorkers = max(1, permits-1)
	}
	return index
}

// projectResolution caches the canonical resolver and its resolved document
// IR for one workspace revision. Full diagnostics can be requested several
// times for the same coherent snapshot; rebuilding TypeLib enum candidates
// and resolving every document on each request is unnecessary work.
func (s *Server) projectResolution(ctx context.Context, project intel.ProjectAnalysisSnapshot, complete bool) (procedureir.Resolver, map[string]procedureir.DocumentIR, map[string]procedureir.DocumentIR) {
	result := s.resolutionCache.getOrBuild(project.Revision, complete, func() projectResolutionResult {
		finishCapability := analysisstats.MeasureCapabilityBuild(ctx, analysisstats.CapabilityResolutionBuildsCounter)
		resolverMeasurement := s.performance.start("project/preparation", performanceStageProjectResolver, "interactive", "")
		s.performance.addCounter(performanceCounterResolutionResolverBuilds, 1, "project/preparation", performanceStageProjectResolver, "interactive", "")
		resolver := workspaceResolutionResolverWithTypeLib(project, complete, s.resolutionTypeLibSymbols)
		resolverMeasurement.finish(len(project.Documents), 0, nil)
		procedureResolver := procedureir.ProcedureOnlyResolver(resolver)
		viewMeasurement := s.performance.start("project/preparation", performanceStageProjectResolutionView, "interactive", "")
		s.performance.addCounter(performanceCounterResolutionViewBuilds, 1, "project/preparation", performanceStageProjectResolutionView, "interactive", "")
		resolved := make(map[string]procedureir.DocumentIR, len(project.Documents))
		diagnosticResolved := make(map[string]procedureir.DocumentIR, len(project.Documents))
		viewMeasurement.finish(len(project.Documents), 0, nil)
		materializationMeasurement := s.performance.start("project/preparation", performanceStageResolutionMaterialize, "interactive", "")
		materializations := 0
		for _, document := range project.Documents {
			// The workspace index resolver intentionally has no TypeDB dependency;
			// apply the cached TypeDB-aware resolver here once per project revision.
			resolved[symbolFileKey(document.IR.Path)] = procedureir.Resolve(document.IR, procedureResolver)
			materializations++
			diagnosticResolved[symbolFileKey(document.IR.Path)] = procedureir.Resolve(document.IR, resolver)
			materializations++
		}
		materializationMeasurement.finish(materializations, 0, nil)
		s.performance.addCounter(performanceCounterResolutionMaterializations, uint64(materializations), "project/preparation", performanceStageResolutionMaterialize, "interactive", "")
		finishCapability(nil)
		return projectResolutionResult{resolver: resolver, resolved: resolved, diagnosticResolved: diagnosticResolved}
	})
	return result.resolver, result.resolved, result.diagnosticResolved
}

func workspaceResolutionResolverWithTypeLib(project intel.ProjectAnalysisSnapshot, complete bool, typeLibSymbols []procedureir.ResolverSymbol) procedureir.Resolver {
	resolverSymbols := make([]procedureir.ResolverSymbol, 0)
	for _, document := range project.Documents {
		module := strings.TrimSpace(document.IR.ModuleName)
		if module == "" {
			module = strings.TrimSuffix(filepath.Base(document.IR.Path), filepath.Ext(document.IR.Path))
		}
		for _, declaration := range document.IR.Declarations {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: declaration.Name, Type: declaration.Type, Module: module, ModuleKind: document.IR.ModuleKind,
				Kind: declaration.Kind, Visibility: declaration.Visibility, File: document.IR.Path,
				Line: declaration.Range.StartLine, Parent: declaration.Parent, IsArray: declaration.IsArray,
				IsConst: declaration.IsConst, ValueShape: declaration.ValueShape,
				Recovered: declaration.Recovered, ConditionalBranches: append([]procedureir.ConditionalBranch(nil), declaration.ConditionalBranches...),
			})
		}
		for _, procedure := range document.IR.Procedures {
			resolverSymbols = append(resolverSymbols, procedureir.ResolverSymbol{
				Name: procedure.Symbol.Name, Type: procedure.Symbol.ReturnType, Module: module, ModuleKind: document.IR.ModuleKind,
				Kind: string(procedure.Symbol.Kind), Visibility: procedure.Symbol.Visibility, File: document.IR.Path,
				Line: procedure.Symbol.DeclarationRange.StartLine, Recovered: procedure.Symbol.Recovered,
				IsArray: procedure.Symbol.IsArray, ValueShape: procedure.Symbol.ValueShape,
				ConditionalBranches: append([]procedureir.ConditionalBranch(nil), procedure.Symbol.ConditionalBranches...),
			})
		}
	}
	resolverSymbols = append(resolverSymbols, typeLibSymbols...)
	return procedureir.NewResolverWithCompleteness(resolverSymbols, complete)
}

func workspaceTypeLibResolverSymbols(typeDB *vbadb.DB) []procedureir.ResolverSymbol {
	if typeDB == nil {
		return nil
	}
	constants := typeDB.AllConstantsList()
	out := make([]procedureir.ResolverSymbol, 0, len(constants))
	for _, constant := range constants {
		if strings.TrimSpace(constant.EnumGroup) == "" {
			continue
		}
		out = append(out, procedureir.ResolverSymbol{
			Name: constant.Name, Parent: constant.EnumGroup, Module: constant.Library,
			ModuleKind: "external", Kind: "enum_member", Visibility: "Public", File: "<typelib>" + constant.Library,
			IsConst: true, ValueShape: procedureir.ValueShapeScalar,
		})
	}
	return out
}

func (s *Server) parseIndexedFileContext(ctx context.Context, file symbols.SourceFile, body []byte) (indexedFileAnalysis, error) {
	release, _, err := s.acquireAnalysisPermit(ctx, "background")
	if err != nil {
		return indexedFileAnalysis{}, err
	}
	defer release()
	snapshot := intel.NewAnalysisSnapshot(intel.Document{
		Path:       file.Path,
		Source:     string(body),
		ModuleKind: file.ModuleKind,
	})
	doc := snapshot.Document()
	defer snapshot.Retire()
	return s.analyzeIndexedDocumentContextClass(ctx, doc, "background")
}

func (s *Server) analyzeIndexedDocumentContextClass(ctx context.Context, doc intel.Document, class string) (indexedFileAnalysis, error) {
	snapshot := doc.Snapshot
	if snapshot == nil || !snapshot.Matches(doc) {
		snapshot = intel.NewAnalysisSnapshot(doc)
		doc = snapshot.Document()
		defer snapshot.Retire()
	}
	syms, err := s.analyzer.DocumentSymbolsContext(ctx, doc)
	if err != nil {
		return indexedFileAnalysis{}, err
	}
	if err := ctx.Err(); err != nil {
		return indexedFileAnalysis{}, err
	}
	s.notifyPerformanceHook("semantic-start", doc.Path)
	semanticMeasurement := s.performance.start("workspace/file", performanceStageSemanticIndexing, class, doc.Path)
	procedureIR, _, err := snapshot.ProcedureIRContext(ctx, func(loadCtx context.Context) (procedureir.DocumentIR, error) {
		parsed, err := snapshot.ParsedDocument()
		if err != nil {
			return procedureir.DocumentIR{}, err
		}
		return procedureir.BuildParsedContext(loadCtx, procedureir.BuildOptions{
			RootDir:    s.opts.RootDir,
			Path:       doc.Path,
			ModuleKind: doc.ModuleKind,
		}, parsed)
	})
	if err != nil {
		semanticMeasurement.finish(0, 0, err)
		return indexedFileAnalysis{}, err
	}
	if err := ctx.Err(); err != nil {
		semanticMeasurement.finish(0, 0, err)
		return indexedFileAnalysis{}, err
	}
	rawCalls, _, err := snapshot.RawCallSites(func() (calls.FileResult, error) {
		return calls.ExtractIR(procedureIR), nil
	})
	if err != nil {
		semanticMeasurement.finish(0, 0, err)
		return indexedFileAnalysis{}, err
	}
	controlFlow, _, err := snapshot.ControlFlowGraphsContext(ctx, func(loadCtx context.Context) (vbacfg.Document, error) {
		return vbacfg.BuildDocumentContext(loadCtx, procedureIR)
	})
	if err != nil {
		semanticMeasurement.finish(0, 0, err)
		return indexedFileAnalysis{}, err
	}
	result := indexedFileAnalysis{
		path:           doc.Path,
		version:        documentVersion(doc),
		moduleKind:     doc.ModuleKind,
		source:         doc.Source,
		symbols:        syms,
		callSites:      rawCalls.CallSites,
		typeReferences: rawCalls.TypeReferences,
		procedureIR:    procedureir.Clone(procedureIR),
		controlFlow:    vbacfg.CloneDocument(controlFlow),
	}
	semanticMeasurement.finish(len(procedureIR.Procedures), 0, nil)
	s.performance.addCounter(performanceCounterWorkspaceSemanticBuilds, 1, "workspace/file", performanceStageSemanticIndexing, class, doc.Path)
	s.notifyPerformanceHook("semantic-ready", doc.Path)
	return result, nil
}

// parseIndexedDeclarationsContext builds only the declaration view needed by
// interactive workspace queries.  It intentionally stops before ProcedureIR,
// call-site, and CFG construction; the semantic phase reparses through the
// existing full callback after declaration readiness is published.
func (s *Server) parseIndexedDeclarationsContext(ctx context.Context, file symbols.SourceFile, body []byte) (indexedFileAnalysis, error) {
	release, _, err := s.acquireAnalysisPermit(ctx, "background")
	if err != nil {
		return indexedFileAnalysis{}, err
	}
	defer release()
	doc := intel.Document{Path: file.Path, Source: string(body), ModuleKind: file.ModuleKind}
	return s.analyzeIndexedDeclarationsDocumentContext(ctx, doc, "background")
}

func (s *Server) analyzeIndexedDeclarationsDocumentContext(ctx context.Context, doc intel.Document, class string) (indexedFileAnalysis, error) {
	snapshot := doc.Snapshot
	if snapshot == nil || !snapshot.Matches(doc) {
		snapshot = intel.NewAnalysisSnapshot(doc)
		defer snapshot.Retire()
	}
	doc = snapshot.Document()
	s.notifyPerformanceHook("declaration-start", doc.Path)
	declarationMeasurement := s.performance.start("workspace/file", performanceStageDeclarationIndexing, class, doc.Path)
	syms, err := s.analyzer.DocumentSymbolsContext(ctx, doc)
	if err != nil {
		declarationMeasurement.finish(len(syms), 0, err)
		return indexedFileAnalysis{}, err
	}
	if err := ctx.Err(); err != nil {
		declarationMeasurement.finish(len(syms), 0, err)
		return indexedFileAnalysis{}, err
	}
	declarationMeasurement.finish(len(syms), 0, nil)
	s.performance.addCounter(performanceCounterWorkspaceDeclarationBuilds, 1, "workspace/file", performanceStageDeclarationIndexing, class, doc.Path)
	s.notifyPerformanceHook("declaration-ready", doc.Path)
	incomplete := false
	parsed, parseErr := snapshot.ParsedDocument()
	if parseErr != nil {
		incomplete = true
	} else {
		_ = parsed.Read(func(view vbaast.ParsedView) error {
			incomplete = view.HasError || view.HasMissing
			return nil
		})
	}
	return indexedFileAnalysis{
		path: doc.Path, moduleKind: doc.ModuleKind, symbols: syms,
		declarationIncomplete: incomplete,
	}, nil
}

func (s *Server) notifyPerformanceHook(stage, path string) {
	if s.performanceHook != nil {
		s.performanceHook(stage, path)
	}
}

func (s *Server) analyzeWorkspaceOverlay(ctx context.Context, doc intel.Document) (indexedFileAnalysis, bool, error) {
	file, included, err := symbols.SourceFileForPath(s.opts.RootDir, s.opts.Config, doc.Path)
	if err != nil || !included {
		return indexedFileAnalysis{}, included, err
	}
	if err := ctx.Err(); err != nil {
		return indexedFileAnalysis{}, true, err
	}
	doc.Path = file.Path
	doc.ModuleKind = file.ModuleKind
	// analyzeIndexedDocument intentionally constructs symbols before IR and
	// calls, so interactive document-local handlers never wait on this permit.
	analysis, err := s.analyzeIndexedDocumentContext(ctx, doc)
	if err != nil {
		return indexedFileAnalysis{}, true, err
	}
	if err := ctx.Err(); err != nil {
		return indexedFileAnalysis{}, true, err
	}
	return analysis, true, nil
}

func (s *Server) analyzeWorkspaceOverlayDeclarations(ctx context.Context, doc intel.Document) (indexedFileAnalysis, intel.Document, bool, error) {
	file, included, err := symbols.SourceFileForPath(s.opts.RootDir, s.opts.Config, doc.Path)
	if err != nil || !included {
		return indexedFileAnalysis{}, doc, included, err
	}
	if err := ctx.Err(); err != nil {
		return indexedFileAnalysis{}, doc, true, err
	}
	doc.Path = file.Path
	doc.ModuleKind = file.ModuleKind
	snapshot := doc.Snapshot
	if snapshot == nil || !snapshot.Matches(doc) {
		snapshot = intel.NewAnalysisSnapshot(doc)
		doc = snapshot.Document()
	}
	analysis, err := s.analyzeIndexedDeclarationsDocumentContext(ctx, doc, "interactive")
	if err != nil {
		return indexedFileAnalysis{}, doc, true, err
	}
	if err := ctx.Err(); err != nil {
		return indexedFileAnalysis{}, doc, true, err
	}
	return analysis, doc, true, nil
}

func (s *Server) analyzeIndexedDocumentContext(ctx context.Context, doc intel.Document) (indexedFileAnalysis, error) {
	return s.analyzeIndexedDocumentContextClass(ctx, doc, "interactive")
}

// diskProcedureSignatures captures the saved declaration baseline in the
// background worker when the initial workspace index has not published this
// path yet. This lets an unsaved first-open rename/removal refresh callers
// without making didOpen parse the disk file synchronously.
func (s *Server) diskProcedureSignatures(ctx context.Context, path string) (map[string]procedureSignature, error) {
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	file, included, err := symbols.SourceFileForPath(s.opts.RootDir, s.opts.Config, path)
	if err != nil || !included {
		return nil, err
	}
	body, err := os.ReadFile(file.Path)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, err
	}
	snapshot := intel.NewAnalysisSnapshot(intel.Document{Path: file.Path, Source: string(body), ModuleKind: file.ModuleKind})
	defer snapshot.Retire()
	syms, err := s.analyzer.DocumentSymbols(snapshot.Document())
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	return procedureSignaturesFromSymbols(syms), nil
}

func (s *Server) cachedDocumentSourceSymbols(doc intel.Document, load intel.DocumentSymbolLoader) ([]intel.Symbol, error) {
	started := time.Now()
	var syms []intel.Symbol
	var hit bool
	var err error
	if doc.Snapshot != nil && doc.Snapshot.Matches(doc) {
		syms, hit, err = doc.Snapshot.SourceSymbols(load)
	} else {
		syms, err = load()
	}
	s.logDocumentCachePerformance("documentSymbols/cache", cacheStatus(hit), doc, len(syms), started, err)
	return syms, err
}

func (s *Server) cachedWorkspaceSymbols(open []intel.Document, query string) ([]intel.Symbol, error) {
	return s.cachedWorkspaceSymbolQuery(open, intel.WorkspaceSymbolQuery{Text: query, Mode: intel.WorkspaceSymbolQueryContains})
}

func (s *Server) cachedWorkspaceSymbolsSnapshot(open []intel.Document) ([]intel.Symbol, error) {
	indexed, err := s.analysis.symbolSnapshot()
	if err != nil {
		return nil, err
	}
	openKeys := make(map[string]bool, len(open)*2)
	for _, doc := range open {
		for _, key := range workspaceSymbolPathKeys(s.opts.RootDir, doc.Path) {
			openKeys[key] = true
		}
	}
	out := indexed[:0]
	for _, sym := range indexed {
		if !hasWorkspaceSymbolPathKey(openKeys, workspaceSymbolPathKeys(s.opts.RootDir, sym.File)) {
			out = append(out, sym)
		}
	}
	for _, doc := range open {
		if s.documentKind(doc) != DocumentKindVBA {
			continue
		}
		syms, symbolErr := s.analyzer.DocumentSymbols(doc)
		if symbolErr == nil {
			out = append(out, syms...)
		}
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
		return out[i].Name < out[j].Name
	})
	return out, nil
}

func (s *Server) cachedWorkspaceSymbolQuery(open []intel.Document, query intel.WorkspaceSymbolQuery) ([]intel.Symbol, error) {
	// Open-document overlays are produced only by the lifecycle pipeline.
	// Queries never synchronously publish them; current snapshot symbols are
	// merged here so interactive handlers remain available while publication is
	// pending and stale disk/overlay entries stay hidden.
	indexed, err := s.queryWorkspaceSymbolIndex(query)
	if err != nil {
		return nil, err
	}
	openKeys := make(map[string]bool, len(open)*2)
	for _, doc := range open {
		for _, key := range workspaceSymbolPathKeys(s.opts.RootDir, doc.Path) {
			openKeys[key] = true
		}
	}
	out := indexed[:0]
	for _, sym := range indexed {
		if hasWorkspaceSymbolPathKey(openKeys, workspaceSymbolPathKeys(s.opts.RootDir, sym.File)) {
			continue
		}
		out = append(out, sym)
	}
	for _, doc := range open {
		if s.documentKind(doc) != DocumentKindVBA {
			continue
		}
		syms, handled := s.analyzer.LightweightDocumentSymbols(doc, query)
		var symbolErr error
		if !handled {
			syms, symbolErr = s.analyzer.DocumentSymbols(doc)
		}
		if symbolErr != nil {
			continue
		}
		for _, sym := range syms {
			if workspaceSymbolMatchesQuery(sym, query) {
				out = append(out, sym)
			}
		}
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
		return out[i].Name < out[j].Name
	})
	return out, nil
}

func (s *Server) queryWorkspaceSymbolIndex(query intel.WorkspaceSymbolQuery) ([]intel.Symbol, error) {
	switch query.Mode {
	case intel.WorkspaceSymbolQueryExact:
		return s.analysis.searchExact(query.Text)
	case intel.WorkspaceSymbolQueryPrefix:
		return s.analysis.searchPrefix(query.Text)
	case intel.WorkspaceSymbolQueryQualified:
		return s.analysis.searchQualified(query.Text)
	case intel.WorkspaceSymbolQueryModule:
		return s.analysis.searchModule(query.Text)
	case intel.WorkspaceSymbolQueryKind:
		return s.analysis.searchKind(query.Text)
	default:
		return s.analysis.searchContains(query.Text)
	}
}

func workspaceSymbolMatchesQuery(sym intel.Symbol, query intel.WorkspaceSymbolQuery) bool {
	text := normalizeSymbolQuery(query.Text)
	name := normalizeSymbolQuery(sym.Name)
	qualified := normalizeSymbolQuery(qualifiedSymbolName(sym))
	switch query.Mode {
	case intel.WorkspaceSymbolQueryExact:
		return name == text
	case intel.WorkspaceSymbolQueryPrefix:
		return strings.HasPrefix(name, text) || strings.HasPrefix(qualified, text)
	case intel.WorkspaceSymbolQueryQualified:
		return qualified == text
	case intel.WorkspaceSymbolQueryModule:
		return normalizeSymbolQuery(sym.Module) == text
	case intel.WorkspaceSymbolQueryKind:
		return normalizeSymbolQuery(sym.Kind) == text
	default:
		return strings.Contains(name, text) || strings.Contains(qualified, text)
	}
}

func workspaceSymbolPathKeys(root, path string) []string {
	if strings.TrimSpace(path) == "" {
		return nil
	}
	keys := []string{symbolFileKey(path)}
	if filepath.IsAbs(path) {
		if rel, err := filepath.Rel(root, path); err == nil {
			keys = append(keys, symbolFileKey(rel))
		}
	} else {
		keys = append(keys, symbolFileKey(filepath.Join(root, filepath.FromSlash(path))))
	}
	return keys
}

func hasWorkspaceSymbolPathKey(set map[string]bool, keys []string) bool {
	for _, key := range keys {
		if key != "" && set[key] {
			return true
		}
	}
	return false
}

func documentSymbolKey(doc intel.Document) string {
	if doc.Path != "" {
		return symbolFileKey(doc.Path)
	}
	if doc.URI != "" {
		if path, err := fileURIToPath(doc.URI); err == nil {
			return symbolFileKey(path)
		}
		return strings.ToLower(doc.URI)
	}
	return ""
}

func symbolFileKey(path string) string {
	if path == "" {
		return ""
	}
	if strings.HasPrefix(path, "file:") {
		if decoded, err := fileURIToPath(path); err == nil {
			path = decoded
		}
	}
	return normalizePathKey(path)
}

type documents struct {
	root          string
	cfg           config.Config
	formsRoot     string
	workbookRoot  string
	readFile      func(string) ([]byte, error)
	beforePublish func()
	mu            sync.RWMutex
	docs          map[string]documentEntry
	keys          map[string]string
	generations   map[string]uint64
	closed        bool
}

type documentEntry struct {
	snapshot   *intel.AnalysisSnapshot
	document   intel.Document
	kind       DocumentKind
	lineIndex  *lineOffsetIndex
	open       bool
	generation uint64
	lifecycle  uint64
}

func newDocuments(root string, sourceRoots ...string) *documents {
	formsRoot := ""
	workbookRoot := ""
	if len(sourceRoots) > 0 {
		formsRoot = sourceRoots[0]
	}
	if len(sourceRoots) > 1 {
		workbookRoot = sourceRoots[1]
	}
	return &documents{
		root: root, formsRoot: formsRoot, workbookRoot: workbookRoot, readFile: os.ReadFile,
		docs: map[string]documentEntry{}, keys: map[string]string{}, generations: map[string]uint64{},
	}
}

func (d *documents) documentKind(doc intel.Document) DocumentKind {
	return DetectDocumentKind(d.root, d.formsRoot, doc.Path, doc.Source)
}

func (entry documentEntry) currentDocument() intel.Document {
	if entry.snapshot != nil {
		return entry.snapshot.Document()
	}
	return entry.document
}

func (entry documentEntry) currentVersion() int32 {
	if entry.snapshot != nil {
		return entry.snapshot.Version()
	}
	return entry.document.Version
}

func (d *documents) nextGenerationLocked(key string) uint64 {
	d.generations[key]++
	return d.generations[key]
}

func (d *documents) open(uri, text string, versions ...int32) (intel.Document, error) {
	return d.openWithIndex(uri, text, newLineOffsetIndex(text), versions...)
}

func (d *documents) openWithIndex(uri, text string, lineIndex *lineOffsetIndex, versions ...int32) (intel.Document, error) {
	doc, err := d.docFromURI(uri, text)
	if err != nil {
		return intel.Document{}, err
	}
	if len(versions) > 0 {
		doc.Version = versions[0]
	}
	kind := d.documentKind(doc)
	if kind != DocumentKindVBA {
		d.mu.Lock()
		if d.closed {
			d.mu.Unlock()
			return intel.Document{}, errDocumentsClosed
		}
		key := normalizePathKey(doc.Path)
		previous := d.docs[key].snapshot
		generation := d.nextGenerationLocked(key)
		d.docs[key] = documentEntry{document: doc, kind: kind, lineIndex: lineIndex, open: true, generation: generation, lifecycle: generation}
		d.keys[uri] = key
		d.mu.Unlock()
		if previous != nil {
			previous.Retire()
		}
		return doc, nil
	}
	snapshot := intel.NewAnalysisSnapshot(doc)
	if d.beforePublish != nil {
		d.beforePublish()
	}
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		snapshot.Retire()
		return intel.Document{}, errDocumentsClosed
	}
	key := normalizePathKey(doc.Path)
	previous := d.docs[key].snapshot
	generation := d.nextGenerationLocked(key)
	d.docs[key] = documentEntry{snapshot: snapshot, document: snapshot.Document(), kind: kind, lineIndex: lineIndex, open: true, generation: generation, lifecycle: generation}
	d.keys[uri] = key
	d.mu.Unlock()
	if previous != nil {
		previous.Retire()
	}
	return snapshot.Document(), nil
}

func (d *documents) change(uri, text string, versions ...int32) (intel.Document, error) {
	doc, _, err := d.changeWithIndex(uri, text, newLineOffsetIndex(text), versions...)
	return doc, err
}

func (d *documents) changeWithIndex(uri, text string, lineIndex *lineOffsetIndex, versions ...int32) (intel.Document, bool, error) {
	d.mu.RLock()
	if d.closed {
		d.mu.RUnlock()
		return intel.Document{}, false, errDocumentsClosed
	}
	key := d.keys[uri]
	if key == "" {
		if path, err := fileURIToPath(uri); err == nil {
			key = normalizePathKey(path)
		}
	}
	entry, ok := d.docs[key]
	d.mu.RUnlock()
	if !ok || !entry.open {
		doc, err := d.openWithIndex(uri, text, lineIndex, versions...)
		return doc, err == nil, err
	}
	if entry.kind != DocumentKindVBA {
		if len(versions) > 0 && versions[0] <= entry.currentVersion() {
			return entry.currentDocument(), false, nil
		}
		doc, err := d.docFromURI(uri, text)
		if err != nil {
			return intel.Document{}, false, err
		}
		if len(versions) > 0 {
			doc.Version = versions[0]
		}
		kind := d.documentKind(doc)
		d.mu.Lock()
		if d.closed {
			d.mu.Unlock()
			return intel.Document{}, false, errDocumentsClosed
		}
		latest := d.docs[key]
		if !latest.open || latest.generation != entry.generation || latest.lifecycle != entry.lifecycle {
			d.mu.Unlock()
			return latest.currentDocument(), false, nil
		}
		generation := d.nextGenerationLocked(key)
		d.docs[key] = documentEntry{document: doc, kind: kind, lineIndex: lineIndex, open: true, generation: generation, lifecycle: entry.lifecycle}
		d.keys[uri] = key
		d.mu.Unlock()
		return doc, true, nil
	}
	current := entry.snapshot.Document()
	current.Source = text
	current.Snapshot = nil
	if len(versions) > 0 {
		current.Version = versions[0]
	}
	snapshot := intel.NewAnalysisSnapshot(current)
	if d.beforePublish != nil {
		d.beforePublish()
	}
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		snapshot.Retire()
		return intel.Document{}, false, errDocumentsClosed
	}
	latest := d.docs[key]
	if latest.snapshot != entry.snapshot || latest.generation != entry.generation || !latest.open {
		if latest.open && latest.snapshot != nil {
			d.keys[uri] = key
			if latest.lifecycle == entry.lifecycle && snapshot.Version() > latest.snapshot.Version() {
				generation := d.nextGenerationLocked(key)
				d.docs[key] = documentEntry{
					snapshot: snapshot, document: snapshot.Document(), kind: DocumentKindVBA,
					lineIndex: lineIndex, open: true, generation: generation, lifecycle: entry.lifecycle,
				}
				d.mu.Unlock()
				latest.snapshot.Retire()
				return snapshot.Document(), true, nil
			}
		}
		d.mu.Unlock()
		snapshot.Retire()
		if latest.open && latest.snapshot != nil {
			return latest.snapshot.Document(), false, nil
		}
		return intel.Document{}, false, errDocumentChangedConcurrently
	}
	generation := d.nextGenerationLocked(key)
	d.docs[key] = documentEntry{snapshot: snapshot, document: snapshot.Document(), kind: DocumentKindVBA, lineIndex: lineIndex, open: true, generation: generation, lifecycle: entry.lifecycle}
	d.keys[uri] = key
	d.mu.Unlock()
	entry.snapshot.Retire()
	return snapshot.Document(), true, nil
}

// applyChanges applies an ordered didChange notification to an open document.
// Ranged changes need a retained source; a full replacement may recover an
// unseen document just as the historic full-sync path did.
func (d *documents) applyChanges(uri string, changes []documentContentChange, version int32) (intel.Document, bool, error) {
	result, err := d.applyChangesWithResult(uri, changes, version)
	return result.document, result.applied, err
}

type documentChangeResult struct {
	document       intel.Document
	applied        bool
	parseMode      string
	fallbackReason string
}

func (d *documents) applyChangesWithResult(uri string, changes []documentContentChange, version int32) (documentChangeResult, error) {
	d.mu.RLock()
	if d.closed {
		d.mu.RUnlock()
		return documentChangeResult{}, errDocumentsClosed
	}
	key := d.keys[uri]
	if key == "" {
		if path, err := fileURIToPath(uri); err == nil {
			key = normalizePathKey(path)
		}
	}
	entry, exists := d.docs[key]
	d.mu.RUnlock()
	if !exists || !entry.open {
		if len(changes) == 0 || changes[0].rng != nil {
			return documentChangeResult{parseMode: "retained", fallbackReason: "document_not_open"}, nil
		}
		source, index, _, _, ok := prepareDocumentContentChanges("", newLineOffsetIndex(""), changes)
		if !ok {
			return documentChangeResult{parseMode: "retained", fallbackReason: "edit_coordinates_unreconciled"}, nil
		}
		doc, err := d.openWithIndex(uri, source, index, version)
		if err != nil {
			return documentChangeResult{}, err
		}
		return documentChangeResult{document: doc, applied: true, parseMode: "full_fallback", fallbackReason: "no_previous_tree"}, nil
	}
	if entry.kind != DocumentKindVBA {
		if version <= entry.currentVersion() {
			return documentChangeResult{document: entry.currentDocument(), parseMode: "retained", fallbackReason: "invalid_version"}, nil
		}
		source, index, _, _, ok := prepareDocumentContentChanges(entry.document.Source, entry.lineIndex, changes)
		if !ok {
			return documentChangeResult{document: entry.currentDocument(), parseMode: "retained", fallbackReason: "edit_coordinates_unreconciled"}, nil
		}
		doc, err := d.docFromURI(uri, source)
		if err != nil {
			return documentChangeResult{}, err
		}
		doc.Version = version
		kind := d.documentKind(doc)
		d.mu.Lock()
		if d.closed {
			d.mu.Unlock()
			return documentChangeResult{}, errDocumentsClosed
		}
		latest := d.docs[key]
		if !latest.open || latest.generation != entry.generation || latest.lifecycle != entry.lifecycle {
			d.mu.Unlock()
			return documentChangeResult{document: latest.currentDocument(), parseMode: "retained", fallbackReason: "document_changed_concurrently"}, nil
		}
		generation := d.nextGenerationLocked(key)
		d.docs[key] = documentEntry{document: doc, kind: kind, lineIndex: index, open: true, generation: generation, lifecycle: entry.lifecycle}
		d.keys[uri] = key
		d.mu.Unlock()
		return documentChangeResult{document: doc, applied: true, parseMode: "full_fallback", fallbackReason: "not_vba_document"}, nil
	}
	if version <= entry.snapshot.Version() {
		return documentChangeResult{document: entry.snapshot.Document(), parseMode: "retained", fallbackReason: "invalid_version"}, nil
	}
	source, index, edits, canIncrementallyParse, ok := prepareDocumentContentChanges(entry.snapshot.Source(), entry.lineIndex, changes)
	if !ok {
		return documentChangeResult{document: entry.snapshot.Document(), parseMode: "retained", fallbackReason: "edit_coordinates_unreconciled"}, nil
	}
	doc := entry.snapshot.Document()
	doc.Source = source
	doc.Version = version
	doc.Snapshot = nil

	var (
		snapshot       *intel.AnalysisSnapshot
		parseMode      = "incremental"
		fallbackReason string
		err            error
	)
	if canIncrementallyParse {
		snapshot, err = intel.NewIncrementalAnalysisSnapshot(doc, entry.snapshot, edits)
	}
	if snapshot == nil {
		parseMode = "full_fallback"
		if !canIncrementallyParse {
			fallbackReason = "full_document_change"
		} else {
			fallbackReason = "incremental_parse_unavailable"
		}
		snapshot, err = fullyParsedSnapshot(doc, entry.snapshot)
	}
	if err != nil || snapshot == nil {
		return documentChangeResult{document: entry.snapshot.Document(), parseMode: "retained", fallbackReason: "full_parse_failed"}, nil
	}
	return d.publishChangedSnapshot(uri, key, entry, snapshot, index, parseMode, fallbackReason)
}

func fullyParsedSnapshot(doc intel.Document, previous *intel.AnalysisSnapshot) (*intel.AnalysisSnapshot, error) {
	parsed, err := vbaast.ParseDocument(doc.Path, []byte(doc.Source))
	if err != nil {
		return nil, err
	}
	snapshot := intel.NewSuccessorAnalysisSnapshotWithParsedDocument(doc, parsed, previous)
	if _, err := snapshot.ParsedDocument(); err != nil {
		snapshot.Retire()
		return nil, err
	}
	return snapshot, nil
}

func (d *documents) publishChangedSnapshot(uri, key string, entry documentEntry, snapshot *intel.AnalysisSnapshot, lineIndex *lineOffsetIndex, parseMode, fallbackReason string) (documentChangeResult, error) {
	d.mu.Lock()
	if d.closed {
		d.mu.Unlock()
		snapshot.Retire()
		return documentChangeResult{}, errDocumentsClosed
	}
	latest, exists := d.docs[key]
	if !exists || !latest.open || latest.snapshot != entry.snapshot || latest.generation != entry.generation || latest.lifecycle != entry.lifecycle {
		d.mu.Unlock()
		snapshot.Retire()
		if latest.open && latest.snapshot != nil {
			return documentChangeResult{document: latest.snapshot.Document(), parseMode: "retained", fallbackReason: "document_changed_concurrently"}, nil
		}
		return documentChangeResult{}, errDocumentChangedConcurrently
	}
	generation := d.nextGenerationLocked(key)
	d.docs[key] = documentEntry{snapshot: snapshot, document: snapshot.Document(), kind: DocumentKindVBA, lineIndex: lineIndex, open: true, generation: generation, lifecycle: entry.lifecycle}
	d.keys[uri] = key
	d.mu.Unlock()
	entry.snapshot.Retire()
	return documentChangeResult{document: snapshot.Document(), applied: true, parseMode: parseMode, fallbackReason: fallbackReason}, nil
}

func (d *documents) close(uri string) {
	d.mu.Lock()
	var snapshot *intel.AnalysisSnapshot
	key := d.keys[uri]
	if key == "" {
		if path, err := fileURIToPath(uri); err == nil {
			key = normalizePathKey(path)
		}
	}
	if key != "" {
		if entry := d.docs[key]; entry.open {
			snapshot = entry.snapshot
			delete(d.docs, key)
		}
		d.nextGenerationLocked(key)
		delete(d.keys, uri)
	}
	d.mu.Unlock()
	if snapshot != nil {
		snapshot.Retire()
	}
}

// invalidateDisk drops a cached closed-file snapshot after a watcher event.
// Open snapshots remain authoritative until didClose.
func (d *documents) invalidateDisk(path string) {
	key := normalizePathKey(path)
	if key == "" {
		return
	}
	d.mu.Lock()
	entry, ok := d.docs[key]
	if ok && entry.open {
		d.mu.Unlock()
		return
	}
	if ok {
		delete(d.docs, key)
	}
	d.nextGenerationLocked(key)
	d.mu.Unlock()
	if ok && entry.snapshot != nil {
		entry.snapshot.Retire()
	}
}

func (d *documents) getOrRead(uri string) (intel.Document, error) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return intel.Document{}, err
	}
	key := normalizePathKey(path)
	for {
		d.mu.RLock()
		if d.closed {
			d.mu.RUnlock()
			return intel.Document{}, errDocumentsClosed
		}
		if entry, ok := d.docs[key]; ok && entry.open {
			d.mu.RUnlock()
			return entry.currentDocument(), nil
		}
		observedGeneration := d.generations[key]
		d.mu.RUnlock()

		body, err := d.readFile(path)
		if err != nil {
			return intel.Document{}, err
		}
		doc := intel.Document{URI: uri, Path: path, Source: string(body), ModuleKind: d.moduleKindForPath(path)}
		kind := d.documentKind(doc)
		if kind != DocumentKindVBA {
			d.mu.Lock()
			if d.closed {
				d.mu.Unlock()
				return intel.Document{}, errDocumentsClosed
			}
			current := d.docs[key]
			if current.open {
				d.mu.Unlock()
				return current.currentDocument(), nil
			}
			if d.generations[key] != observedGeneration {
				d.mu.Unlock()
				continue
			}
			generation := d.nextGenerationLocked(key)
			d.docs[key] = documentEntry{document: doc, kind: kind, generation: generation}
			d.mu.Unlock()
			if current.snapshot != nil {
				current.snapshot.Retire()
			}
			return doc, nil
		}
		candidate := intel.NewAnalysisSnapshot(doc)
		if d.beforePublish != nil {
			d.beforePublish()
		}
		d.mu.Lock()
		if d.closed {
			d.mu.Unlock()
			candidate.Retire()
			return intel.Document{}, errDocumentsClosed
		}
		current := d.docs[key]
		if current.open {
			d.mu.Unlock()
			candidate.Retire()
			return current.currentDocument(), nil
		}
		if current.snapshot != nil && current.snapshot.SourceHash() == candidate.SourceHash() &&
			current.snapshot.URI() == candidate.URI() && current.snapshot.ModuleKind() == candidate.ModuleKind() {
			d.mu.Unlock()
			candidate.Retire()
			return current.snapshot.Document(), nil
		}
		if d.generations[key] != observedGeneration {
			d.mu.Unlock()
			candidate.Retire()
			continue
		}
		generation := d.nextGenerationLocked(key)
		d.docs[key] = documentEntry{snapshot: candidate, document: candidate.Document(), kind: DocumentKindVBA, generation: generation}
		d.mu.Unlock()
		if current.snapshot != nil {
			current.snapshot.Retire()
		}
		return candidate.Document(), nil
	}
}

func (d *documents) openDocuments() []intel.Document {
	d.mu.RLock()
	defer d.mu.RUnlock()
	out := make([]intel.Document, 0, len(d.docs))
	for _, entry := range d.docs {
		if entry.open && entry.kind == DocumentKindVBA && entry.snapshot != nil {
			out = append(out, entry.snapshot.Document())
		}
	}
	return out
}

func (d *documents) isOpen(uri string) bool {
	d.mu.RLock()
	defer d.mu.RUnlock()
	key := d.keys[uri]
	if key == "" {
		if path, err := fileURIToPath(uri); err == nil {
			key = normalizePathKey(path)
		}
	}
	return key != "" && d.docs[key].open
}

func (d *documents) closeAll() {
	d.mu.Lock()
	snapshots := make([]*intel.AnalysisSnapshot, 0, len(d.docs))
	for _, entry := range d.docs {
		if entry.snapshot != nil {
			snapshots = append(snapshots, entry.snapshot)
		}
	}
	d.docs = make(map[string]documentEntry)
	d.keys = make(map[string]string)
	d.closed = true
	d.mu.Unlock()
	for _, snapshot := range snapshots {
		snapshot.Retire()
	}
}

var (
	errDocumentsClosed             = errors.New("LSP document snapshot store is closed")
	errDocumentChangedConcurrently = errors.New("LSP document changed concurrently")
)

func (d *documents) docFromURI(uri, text string) (intel.Document, error) {
	path, err := fileURIToPath(uri)
	if err != nil {
		return intel.Document{}, err
	}
	return intel.Document{URI: uri, Path: path, Source: text, ModuleKind: d.moduleKindForPath(path)}, nil
}

func (d *documents) uriForDisplayPath(path string) string {
	if strings.HasPrefix(path, "file:") {
		return path
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(d.root, filepath.FromSlash(path))
	}
	return pathToFileURI(path)
}

func fileURIToPath(raw string) (string, error) {
	u, err := url.Parse(raw)
	if err != nil {
		return "", err
	}
	if u.Scheme != "file" {
		return "", fmt.Errorf("unsupported URI scheme %q", u.Scheme)
	}
	path, err := url.PathUnescape(u.EscapedPath())
	if err != nil {
		return "", err
	}
	if runtime.GOOS == "windows" {
		if u.Host != "" {
			path = `\\` + u.Host + filepath.FromSlash(path)
		} else {
			path = strings.TrimPrefix(path, "/")
			path = filepath.FromSlash(path)
		}
	} else {
		path = filepath.FromSlash(path)
	}
	return filepath.Clean(path), nil
}

func pathToFileURI(path string) string {
	path = filepath.Clean(path)
	host := ""
	if runtime.GOOS == "windows" {
		vol := filepath.VolumeName(path)
		if strings.HasPrefix(vol, `\\`) {
			rest := strings.TrimPrefix(path, vol)
			hostShare := strings.TrimPrefix(vol, `\\`)
			parts := strings.SplitN(hostShare, `\`, 2)
			if len(parts) == 2 {
				host = parts[0]
				path = "/" + parts[1] + filepath.ToSlash(rest)
			} else {
				path = "/" + filepath.ToSlash(path)
			}
		} else {
			path = "/" + filepath.ToSlash(path)
		}
	} else {
		path = filepath.ToSlash(path)
	}
	return (&url.URL{Scheme: "file", Host: host, Path: path}).String()
}

func normalizePathKey(path string) string {
	clean := filepath.Clean(path)
	if runtime.GOOS == "windows" {
		clean = strings.ToLower(clean)
	}
	return clean
}

func moduleKindForPath(path string) string {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".cls":
		return "class"
	case ".frm":
		return "form"
	default:
		return "standard"
	}
}

func (d *documents) moduleKindForPath(path string) string {
	if file, included, err := symbols.SourceFileForPath(d.root, d.cfg, path); err == nil && included && file.ModuleKind != "" {
		return string(file.ModuleKind)
	}
	if isWorkbookModulePath(d.root, d.workbookRoot, path) {
		return "document"
	}
	return moduleKindForPath(path)
}

func isWorkbookModulePath(root, configuredWorkbookRoot, path string) bool {
	if strings.TrimSpace(root) == "" {
		root = "."
	}
	workbookRoot := strings.TrimSpace(configuredWorkbookRoot)
	if workbookRoot == "" {
		workbookRoot = filepath.Join("src", "workbook")
	}
	if !filepath.IsAbs(workbookRoot) {
		workbookRoot = filepath.Join(root, workbookRoot)
	}
	workbookRoot, err := filepath.Abs(workbookRoot)
	if err != nil {
		return false
	}
	if !filepath.IsAbs(path) {
		path = filepath.Join(root, path)
	}
	path, err = filepath.Abs(path)
	if err != nil {
		return false
	}
	rel, err := filepath.Rel(filepath.Clean(workbookRoot), filepath.Clean(path))
	if err != nil {
		return false
	}
	return rel != ".." && !strings.HasPrefix(rel, ".."+string(filepath.Separator))
}

func documentSupportsFormatting(doc intel.Document) bool {
	ext := strings.ToLower(filepath.Ext(doc.Path))
	return ext == ".bas" || ext == ".cls"
}

func documentIsClass(doc intel.Document) bool {
	return strings.EqualFold(doc.ModuleKind, "class") || strings.EqualFold(filepath.Ext(doc.Path), ".cls")
}

func fullDocumentRange(source string) protocol.Range {
	source = strings.ReplaceAll(source, "\r\n", "\n")
	source = strings.ReplaceAll(source, "\r", "\n")
	lines := strings.Split(source, "\n")
	lastLine := len(lines) - 1
	lastChar := 0
	if lastLine >= 0 {
		lastChar = utf16Len(lines[lastLine])
	}
	return protocol.Range{
		Start: protocol.Position{Line: 0, Character: 0},
		End:   protocol.Position{Line: protocol.UInteger(max(0, lastLine)), Character: protocol.UInteger(max(0, lastChar))},
	}
}

func toProtocolRange(r intel.Range) protocol.Range {
	return protocol.Range{Start: toProtocolPosition(r.Start), End: toProtocolPosition(r.End)}
}

func toProtocolRangePtr(r intel.Range) *protocol.Range {
	out := toProtocolRange(r)
	return &out
}

func toProtocolPosition(pos intel.Position) protocol.Position {
	return protocol.Position{Line: protocol.UInteger(max(0, pos.Line)), Character: protocol.UInteger(max(0, pos.Character))}
}

func fromProtocolPosition(pos protocol.Position) intel.Position {
	return intel.Position{Line: int(pos.Line), Character: int(pos.Character)}
}

func fromProtocolRange(r protocol.Range) intel.Range {
	return intel.Range{Start: fromProtocolPosition(r.Start), End: fromProtocolPosition(r.End)}
}

func utf16Len(s string) int {
	return len(utf16.Encode([]rune(s)))
}

func diagnosticSeverity(severity string) protocol.DiagnosticSeverity {
	switch strings.ToLower(severity) {
	case "error":
		return protocol.DiagnosticSeverityError
	case "information", "info":
		return protocol.DiagnosticSeverityInformation
	case "hint":
		return protocol.DiagnosticSeverityHint
	default:
		return protocol.DiagnosticSeverityWarning
	}
}

func encodeSemanticTokens(tokens []intel.SemanticToken) []protocol.UInteger {
	out := make([]protocol.UInteger, 0, len(tokens)*5)
	prevLine, prevStart := 0, 0
	for _, token := range tokens {
		line := max(0, token.Range.Start.Line)
		start := max(0, token.Range.Start.Character)
		length := max(0, token.Range.End.Character-token.Range.Start.Character)
		if length == 0 || token.Range.End.Line != token.Range.Start.Line {
			continue
		}
		deltaLine := line - prevLine
		deltaStart := start
		if deltaLine == 0 {
			deltaStart = start - prevStart
		}
		if deltaStart < 0 {
			continue
		}
		typeIndex := semanticTokenTypeIndex(token.Type)
		if typeIndex < 0 {
			continue
		}
		out = append(out,
			protocol.UInteger(deltaLine),
			protocol.UInteger(deltaStart),
			protocol.UInteger(length),
			protocol.UInteger(typeIndex),
			protocol.UInteger(semanticTokenModifierMask(token.Modifiers)),
		)
		prevLine = line
		prevStart = start
	}
	return out
}

func semanticTokenTypeIndex(tokenType string) int {
	for i, candidate := range intel.SemanticTokenTypes {
		if candidate == tokenType {
			return i
		}
	}
	return -1
}

func semanticTokenModifierMask(modifiers []string) int {
	mask := 0
	for _, modifier := range modifiers {
		for i, candidate := range intel.SemanticTokenModifiers {
			if modifier == candidate {
				mask |= 1 << i
			}
		}
	}
	return mask
}

func symbolKind(kind string) protocol.SymbolKind {
	switch strings.ToLower(kind) {
	case "module":
		return protocol.SymbolKindModule
	case "class":
		return protocol.SymbolKindClass
	case "sub", "function", "property", "property_get", "property_let", "property_set":
		return protocol.SymbolKindFunction
	case "const":
		return protocol.SymbolKindConstant
	case "field", "module_variable":
		return protocol.SymbolKindField
	case "local_variable":
		return protocol.SymbolKindVariable
	case "enum":
		return protocol.SymbolKindEnum
	case "event":
		return protocol.SymbolKindEvent
	default:
		return protocol.SymbolKindObject
	}
}

func completionItemKind(kind string) protocol.CompletionItemKind {
	switch strings.ToLower(kind) {
	case "method":
		return protocol.CompletionItemKindMethod
	case "function":
		return protocol.CompletionItemKindFunction
	case "property":
		return protocol.CompletionItemKindProperty
	case "variable":
		return protocol.CompletionItemKindVariable
	case "type":
		return protocol.CompletionItemKindClass
	case "constant":
		return protocol.CompletionItemKindConstant
	case "keyword":
		return protocol.CompletionItemKindKeyword
	case "snippet":
		return protocol.CompletionItemKindSnippet
	default:
		return protocol.CompletionItemKindText
	}
}

func parameterLabel(param intel.Parameter) string {
	var b strings.Builder
	if param.Optional {
		b.WriteString("Optional ")
	}
	b.WriteString(param.Name)
	if param.Type != "" {
		b.WriteString(" As ")
		b.WriteString(param.Type)
	}
	return b.String()
}

func completionTriggerCharacters() []string {
	return []string{".", "\"", "'", "@"}
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}

type stdioReadWriteCloser struct{}

func (stdioReadWriteCloser) Read(p []byte) (int, error)  { return os.Stdin.Read(p) }
func (stdioReadWriteCloser) Write(p []byte) (int, error) { return os.Stdout.Write(p) }
func (stdioReadWriteCloser) Close() error                { return nil }

type rpcHandler struct {
	handler glsp.Handler
}

func (h rpcHandler) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	params := []byte("{}")
	if req.Params != nil {
		params = *req.Params
	}
	glspCtx := &glsp.Context{
		Method: req.Method,
		Params: params,
		Notify: func(method string, params any) {
			_ = conn.Notify(ctx, method, params)
		},
		Call: func(method string, params any, result any) {
			_ = conn.Call(ctx, method, params, result)
		},
	}
	result, validMethod, validParams, err := h.handler.Handle(glspCtx)
	if !validMethod {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeMethodNotFound, Message: "method not found"})
		}
		return
	}
	if !validParams {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInvalidParams, Message: "invalid params"})
		}
		return
	}
	if err != nil {
		if !req.Notif {
			_ = conn.ReplyWithError(ctx, req.ID, &jsonrpc2.Error{Code: jsonrpc2.CodeInternalError, Message: err.Error()})
		}
		return
	}
	if !req.Notif {
		_ = conn.Reply(ctx, req.ID, result)
	}
	if req.Method == "exit" {
		_ = conn.Close()
	}
}
