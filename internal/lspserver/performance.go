package lspserver

import (
	"context"
	"errors"
	"log"
	"os"
	"sort"
	"strings"
	"sync"
	"time"

	"github.com/harumiWeb/xlflow/internal/vba/analysisstats"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
)

const startupIDEnv = "XLFLOW_LSP_STARTUP_ID"

type startupContext struct {
	id      string
	started time.Time
}

func startupContextFromEnvironment(enabled bool) *startupContext {
	if !enabled {
		return nil
	}
	id := strings.TrimSpace(os.Getenv(startupIDEnv))
	if id == "" {
		return nil
	}
	return &startupContext{id: id, started: time.Now()}
}

// Performance stage names are intentionally stable: they are consumed by
// developer benchmark and profile scripts, but never enter an LSP payload.
const (
	performanceStageWorkspaceDiscovery    = "workspaceDiscovery"
	performanceStageDeclarationIndexing   = "declarationIndexing"
	performanceStageSemanticIndexing      = "semanticIndexing"
	performanceStageProjectSnapshot       = "projectSnapshot"
	performanceStageProjectResolver       = "projectResolver"
	performanceStageProjectResolutionView = "projectResolutionView"
	performanceStageResolutionMaterialize = "projectResolutionMaterialization"
	performanceStageProjectEffects        = "projectEffectSummary"
	performanceStageProjectConstants      = "projectConstants"
	performanceStageProjectCapabilities   = "projectCapabilityPlan"
	performanceStageProjectChange         = "projectChange"
	performanceStageDependencyUpdate      = "dependencyFingerprintUpdate"
	performanceStagePermitWait            = "permitWait"
)

const (
	performanceCounterWorkspaceFilesDiscovered       = "workspace_files_discovered"
	performanceCounterWorkspaceDeclarationBuilds     = "workspace_declaration_builds"
	performanceCounterWorkspaceSemanticBuilds        = "workspace_semantic_builds"
	performanceCounterInteractiveIndexBuilds         = "interactive_index_builds"
	performanceCounterInteractiveIndexHits           = "interactive_index_hits"
	performanceCounterProcedureCatalogBuilds         = "procedure_catalog_builds"
	performanceCounterProcedureCatalogReuses         = "procedure_catalog_reuses"
	performanceCounterInteractiveExactQueries        = "interactive_exact_queries"
	performanceCounterInteractivePrefixQueries       = "interactive_prefix_queries"
	performanceCounterInteractiveQualifiedQueries    = "interactive_qualified_queries"
	performanceCounterFullDocumentSymbolBuilds       = "full_document_symbol_builds"
	performanceCounterInteractiveFullSymbolFallbacks = "interactive_full_symbol_fallbacks"
	performanceCounterProjectSnapshotBuilds          = "project_snapshot_builds"
	performanceCounterResolutionResolverBuilds       = "resolution_resolver_builds"
	performanceCounterResolutionViewBuilds           = "resolution_view_builds"
	performanceCounterCanonicalResolverBuilds        = "canonical_resolver_builds"
	performanceCounterProcedureResolverViews         = "procedure_resolver_views"
	performanceCounterFullResolverViews              = "full_resolver_views"
	performanceCounterResolutionOverlayBuilds        = "resolution_overlay_builds"
	performanceCounterResolutionMaterializations     = "resolution_materializations"
	performanceCounterProcedureFingerprintBuilds     = "procedure_fingerprint_builds"
	performanceCounterProcedureFingerprintReuses     = "procedure_fingerprint_reuses"
	performanceCounterDependencyNodesUpdated         = "dependency_nodes_updated"
	performanceCounterDependencyEdgesUpdated         = "dependency_edges_updated"
	performanceCounterProceduresRevisited            = "procedures_revisited"
	performanceCounterFastDiagnosticRuns             = "fast_diagnostic_runs"
	performanceCounterFullDiagnosticRuns             = "full_diagnostic_runs"
	performanceCounterBackgroundPermitWaits          = "background_permit_waits"
	performanceCounterInteractivePermitWaits         = "interactive_permit_waits"
	performanceCounterProjectCacheHits               = "project_cache_hits"
	performanceCounterProjectCacheMisses             = "project_cache_misses"
	performanceCounterProjectCacheRebuilds           = "project_cache_rebuilds"
	performanceCounterProjectDependencyInvalidations = "project_dependency_invalidations"
	performanceCounterProjectCacheReusedEntries      = "project_cache_reused_entries"
)

var performanceCounterNames = [...]string{
	performanceCounterWorkspaceFilesDiscovered,
	performanceCounterWorkspaceDeclarationBuilds,
	performanceCounterWorkspaceSemanticBuilds,
	performanceCounterInteractiveIndexBuilds,
	performanceCounterInteractiveIndexHits,
	performanceCounterProcedureCatalogBuilds,
	performanceCounterProcedureCatalogReuses,
	performanceCounterInteractiveExactQueries,
	performanceCounterInteractivePrefixQueries,
	performanceCounterInteractiveQualifiedQueries,
	performanceCounterFullDocumentSymbolBuilds,
	performanceCounterInteractiveFullSymbolFallbacks,
	performanceCounterProjectSnapshotBuilds,
	performanceCounterResolutionResolverBuilds,
	performanceCounterResolutionViewBuilds,
	performanceCounterCanonicalResolverBuilds,
	performanceCounterProcedureResolverViews,
	performanceCounterFullResolverViews,
	performanceCounterResolutionOverlayBuilds,
	performanceCounterResolutionMaterializations,
	performanceCounterProcedureFingerprintBuilds,
	performanceCounterProcedureFingerprintReuses,
	performanceCounterDependencyNodesUpdated,
	performanceCounterDependencyEdgesUpdated,
	performanceCounterProceduresRevisited,
	performanceCounterFastDiagnosticRuns,
	performanceCounterFullDiagnosticRuns,
	performanceCounterBackgroundPermitWaits,
	performanceCounterInteractivePermitWaits,
	performanceCounterProjectCacheHits,
	performanceCounterProjectCacheMisses,
	performanceCounterProjectCacheRebuilds,
	performanceCounterProjectDependencyInvalidations,
	performanceCounterProjectCacheReusedEntries,
}

// performanceRecorder is deliberately separate from analysisstats. The
// latter is request-scoped and owned by the analyzer; this recorder covers
// server-lifetime workspace work that has no diagnostic context.
//
// A nil recorder is used when --performance-log is disabled. Callers retain a
// nil-safe pointer and therefore do not pay a clock, lock, or map cost on the
// normal path.
type performanceRecorder struct {
	logger  *log.Logger
	enabled bool
	mu      sync.Mutex
	// startup is only populated for a server process started by the VS Code
	// client with a per-attempt correlation ID. It remains nil for ordinary
	// callers and for performance logging without startup correlation.
	startup  *startupTelemetry
	counters map[string]uint64
}

type startupTelemetry struct {
	logger  *log.Logger
	id      string
	started time.Time
	mu      sync.Mutex
	events  map[string]struct{}
}

func newStartupTelemetry(logger *log.Logger, startup *startupContext) *startupTelemetry {
	if logger == nil || startup == nil || startup.id == "" {
		return nil
	}
	return &startupTelemetry{
		logger: logger, id: startup.id, started: startup.started,
		events: make(map[string]struct{}),
	}
}

func newPerformanceRecorder(enabled bool, logger *log.Logger) *performanceRecorder {
	if !enabled || logger == nil {
		return nil
	}
	counters := make(map[string]uint64, len(performanceCounterNames))
	for _, name := range performanceCounterNames {
		counters[name] = 0
	}
	return &performanceRecorder{logger: logger, enabled: true, counters: counters}
}

func (p *performanceRecorder) setStartup(startup *startupContext) {
	if p == nil || !p.enabled {
		return
	}
	p.startup = newStartupTelemetry(p.logger, startup)
}

func (p *performanceRecorder) startupEvent(event string) {
	if p == nil || !p.enabled || p.startup == nil {
		return
	}
	p.startup.event(event)
}

func (p *performanceRecorder) firstSuccessfulInteractive(operation string, resultCount int) {
	if p == nil || !p.enabled || p.startup == nil || resultCount <= 0 {
		return
	}
	event := ""
	switch operation {
	case "textDocument/hover":
		event = "firstHoverHandled"
	case "textDocument/definition":
		event = "firstDefinitionHandled"
	case "textDocument/completion":
		event = "firstCompletionHandled"
	default:
		return
	}
	p.startup.event(event, resultCount)
}

func (t *startupTelemetry) event(event string, resultCount ...int) {
	if t == nil || t.logger == nil || event == "" {
		return
	}
	t.mu.Lock()
	if _, ok := t.events[event]; ok {
		t.mu.Unlock()
		return
	}
	t.events[event] = struct{}{}
	t.mu.Unlock()
	now := time.Now()
	count := 0
	if len(resultCount) != 0 {
		count = resultCount[0]
	}
	t.logger.Printf(
		"performance operation=%q startup_id=%q event=%q elapsed_ms=%.3f wall_time_unix_ns=%d result_count=%d outcome=%q",
		"lsp/startup", t.id, event,
		float64(now.Sub(t.started))/float64(time.Millisecond), now.UnixNano(), count, "ok",
	)
}

type performanceStageMeasurement struct {
	recorder  *performanceRecorder
	operation string
	stage     string
	class     string
	path      string
	started   time.Time
}

func (p *performanceRecorder) start(operation, stage, class, path string) performanceStageMeasurement {
	if p == nil || !p.enabled {
		return performanceStageMeasurement{}
	}
	return performanceStageMeasurement{recorder: p, operation: operation, stage: stage, class: class, path: path, started: time.Now()}
}

func (m performanceStageMeasurement) finish(resultCount int, wait time.Duration, err error) {
	if m.recorder == nil || m.recorder.logger == nil {
		return
	}
	outcome := "ok"
	if err != nil {
		if errors.Is(err, context.Canceled) {
			outcome = "canceled"
		} else {
			outcome = "error"
		}
	}
	m.recorder.logger.Printf(
		"performance operation=%q stage=%q class=%q path=%q elapsed_ms=%.3f wait_ms=%.3f result_count=%d outcome=%q",
		m.operation, m.stage, m.class, m.path,
		float64(time.Since(m.started))/float64(time.Millisecond),
		float64(wait)/float64(time.Millisecond), resultCount, outcome,
	)
}

func (p *performanceRecorder) addCounter(name string, value uint64, operation, stage, class, path string) {
	if p == nil || !p.enabled || p.logger == nil {
		return
	}
	p.mu.Lock()
	p.counters[name] += value
	total := p.counters[name]
	p.mu.Unlock()
	if value == 0 {
		return
	}
	p.logger.Printf(
		"performance operation=%q stage=%q class=%q path=%q counter=%q value=%d total=%d outcome=%q",
		operation, stage, class, path, name, value, total, "counter",
	)
}

func (p *performanceRecorder) logCounterSnapshot(operation, stage, class, path string) {
	if p == nil || !p.enabled || p.logger == nil {
		return
	}
	p.mu.Lock()
	values := make(map[string]uint64, len(performanceCounterNames))
	for _, name := range performanceCounterNames {
		values[name] = p.counters[name]
	}
	p.mu.Unlock()
	for _, name := range performanceCounterNames {
		p.logger.Printf(
			"performance operation=%q stage=%q class=%q path=%q counter=%q value=%d total=%d outcome=%q",
			operation, stage, class, path, name, 0, values[name], "counter_snapshot",
		)
	}
}

func (p *performanceRecorder) counterTotal(name string) uint64 {
	if p == nil || !p.enabled {
		return 0
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	return p.counters[name]
}

// acquireAnalysisPermit keeps the existing bounded worker budget while making
// contention visible. Interactive waiters are registered before they try the
// channel, and background waiters sleep on a generation signal instead of
// competing in the channel's FIFO-less send queue. This gives a released slot
// to an interactive request before resuming background work.
func (s *Server) acquireAnalysisPermit(ctx context.Context, class string) (func(), time.Duration, error) {
	if ctx == nil {
		ctx = context.Background()
	}
	if err := ctx.Err(); err != nil {
		return nil, 0, err
	}
	if s.analysisPermits == nil {
		return func() {}, 0, nil
	}
	if class == "interactive" {
		s.analysisPermitMu.Lock()
		s.interactivePermitWaiters.Add(1)
		s.analysisPermitMu.Unlock()
		defer s.interactivePermitWaiters.Add(-1)
	}

	var measurement performanceStageMeasurement
	waited := false
	for {
		s.analysisPermitMu.Lock()
		if s.analysisPermitChanged == nil {
			s.analysisPermitChanged = make(chan struct{})
		}
		changed := s.analysisPermitChanged
		backgroundBlocked := class == "background" && s.interactivePermitWaiters.Load() > 0
		if !backgroundBlocked {
			select {
			case s.analysisPermits <- struct{}{}:
				s.analysisPermitMu.Unlock()
				if !waited {
					return s.releaseAnalysisPermit, 0, nil
				}
				wait := time.Duration(0)
				if measurement.recorder != nil {
					wait = time.Since(measurement.started)
				}
				counter := performanceCounterInteractivePermitWaits
				if class == "background" {
					counter = performanceCounterBackgroundPermitWaits
				}
				s.performance.addCounter(counter, 1, "scheduler/permit", performanceStagePermitWait, class, "")
				measurement.finish(0, wait, nil)
				return s.releaseAnalysisPermit, wait, nil
			default:
			}
		}
		s.analysisPermitMu.Unlock()
		if !waited {
			measurement = s.performance.start("scheduler/permit", performanceStagePermitWait, class, "")
			waited = true
		}
		select {
		case <-changed:
		case <-time.After(10 * time.Millisecond):
		case <-ctx.Done():
			wait := time.Duration(0)
			if measurement.recorder != nil {
				wait = time.Since(measurement.started)
			}
			counter := performanceCounterInteractivePermitWaits
			if class == "background" {
				counter = performanceCounterBackgroundPermitWaits
			}
			s.performance.addCounter(counter, 1, "scheduler/permit", performanceStagePermitWait, class, "")
			measurement.finish(0, wait, ctx.Err())
			return nil, wait, ctx.Err()
		}
	}
}

func (s *Server) releaseAnalysisPermit() {
	if s.analysisPermits == nil {
		return
	}
	<-s.analysisPermits
	s.analysisPermitMu.Lock()
	if s.analysisPermitChanged == nil {
		s.analysisPermitChanged = make(chan struct{})
	}
	close(s.analysisPermitChanged)
	s.analysisPermitChanged = make(chan struct{})
	s.analysisPermitMu.Unlock()
}

type performanceMeasurement struct {
	server    *Server
	operation string
	phase     string
	document  intel.Document
	started   time.Time
}

func (s *Server) logDiagnosticStages(doc intel.Document, generation uint64, mode intel.DiagnosticMode, recorder *analysisstats.Recorder) {
	if !s.opts.PerformanceLog || recorder == nil {
		return
	}
	phase := "full"
	if mode == intel.DiagnosticModeFast {
		phase = "fast"
	}
	stages, counters := recorder.Snapshot()
	for _, stage := range stages {
		s.logger.Printf(
			"performance operation=%q uri=%q path=%q version=%d generation=%d phase=%q stage=%q elapsed_ms=%.3f wait_ms=%.3f result_count=%d outcome=%q",
			"diagnostics/stage", doc.URI, doc.Path, doc.Version, generation, phase, stage.Name,
			float64(stage.Elapsed)/float64(time.Millisecond), float64(stage.Wait)/float64(time.Millisecond), stage.ResultCount, stage.Outcome,
		)
	}
	names := make([]string, 0, len(counters))
	for name := range counters {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		value := counters[name]
		s.logger.Printf(
			"performance operation=%q uri=%q path=%q version=%d generation=%d phase=%q stage=%q elapsed_ms=0 wait_ms=0 result_count=%d outcome=%q",
			"diagnostics/stage", doc.URI, doc.Path, doc.Version, generation, phase, name, value, "counter",
		)
	}
}

func (s *Server) startPerformance(operation string, doc intel.Document) *performanceMeasurement {
	measurement := s.startPerformanceURI(operation, doc.URI)
	measurement.setDocument(doc)
	return measurement
}

func (s *Server) startPerformanceURI(operation, uri string) *performanceMeasurement {
	if !s.opts.PerformanceLog {
		return nil
	}
	return &performanceMeasurement{
		server:    s,
		operation: operation,
		document:  intel.Document{URI: uri},
		started:   time.Now(),
	}
}

func (m *performanceMeasurement) setDocument(doc intel.Document) {
	if m != nil {
		m.document = doc
	}
}

func (m *performanceMeasurement) setPhase(phase string) {
	if m != nil {
		m.phase = phase
	}
}

func (m *performanceMeasurement) finish(resultCount int, err error) {
	if m == nil {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	doc := m.document
	m.server.logger.Printf(
		"performance operation=%q stage=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f wait_ms=0 result_count=%d outcome=%q",
		m.operation,
		m.operation,
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(m.started))/float64(time.Millisecond),
		resultCount,
		outcome,
	)
	if err == nil {
		m.server.performance.firstSuccessfulInteractive(m.operation, resultCount)
	}
}

func (m *performanceMeasurement) finishDiagnostics(resultCount int, generation uint64, discarded bool) {
	if m == nil {
		return
	}
	doc := m.document
	outcome := "ok"
	if discarded {
		outcome = "discarded"
	}
	m.server.logger.Printf(
		"performance operation=%q stage=%q phase=%q uri=%q path=%q version=%d generation=%d bytes=%d lines=%d elapsed_ms=%.3f wait_ms=0 result_count=%d outcome=%q discarded=%t",
		m.operation,
		m.operation,
		m.phase,
		doc.URI,
		doc.Path,
		doc.Version,
		generation,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(m.started))/float64(time.Millisecond),
		resultCount,
		outcome,
		discarded,
	)
}

func (s *Server) logWorkspaceOverlayPerformance(doc intel.Document, generation uint64, started time.Time, err error, discarded bool) {
	if !s.opts.PerformanceLog {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	} else if discarded {
		outcome = "discarded"
	}
	s.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d generation=%d bytes=%d lines=%d elapsed_ms=%.3f outcome=%q discarded=%t",
		"workspaceSymbols/overlay", doc.URI, doc.Path, doc.Version, generation,
		len(doc.Source), sourceLineCount(doc.Source),
		float64(time.Since(started))/float64(time.Millisecond), outcome, discarded,
	)
}

func (s *Server) logInitialWorkspaceIndexPerformance(fileCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	if err == nil {
		s.performance.startupEvent("semanticIndexReady")
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q elapsed_ms=%.3f file_count=%d outcome=%q",
		"workspaceSymbols/index/initial",
		float64(time.Since(started))/float64(time.Millisecond),
		fileCount,
		outcome,
	)
	s.performance.logCounterSnapshot("workspaceSymbols/index/initial", performanceStageWorkspaceDiscovery, "background", s.opts.RootDir)
}

func (s *Server) logInitialWorkspaceDeclarationPerformance(fileCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	if err == nil {
		s.performance.startupEvent("declarationIndexReady")
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q elapsed_ms=%.3f file_count=%d outcome=%q",
		"workspaceDeclarations/index/initial",
		float64(time.Since(started))/float64(time.Millisecond),
		fileCount,
		outcome,
	)
	s.performance.logCounterSnapshot("workspaceDeclarations/index/initial", performanceStageWorkspaceDiscovery, "background", s.opts.RootDir)
}

func (s *Server) logDocumentCachePerformance(operation, cache string, doc intel.Document, resultCount int, started time.Time, err error) {
	if !s.opts.PerformanceLog {
		return
	}
	outcome := "ok"
	if err != nil {
		outcome = "error"
	}
	s.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q cache=%q",
		operation,
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(started))/float64(time.Millisecond),
		resultCount,
		outcome,
		cache,
	)
}

func (s *Server) logDocumentChangePerformance(uri string, version int32, change documentChangeResult, started time.Time) {
	if !s.opts.PerformanceLog {
		return
	}
	doc := change.document
	if doc.URI == "" {
		doc.URI = uri
	}
	// A retained change reports the attempted LSP version, even though the
	// source fields intentionally describe the last valid snapshot.
	doc.Version = version
	s.logger.Printf(
		"performance operation=%q uri=%q path=%q version=%d bytes=%d lines=%d elapsed_ms=%.3f result_count=%d outcome=%q parse_mode=%q fallback_reason=%q",
		"textDocument/didChange/parse",
		doc.URI,
		doc.Path,
		doc.Version,
		len(doc.Source),
		sourceLineCount(doc.Source),
		float64(time.Since(started))/float64(time.Millisecond),
		0,
		map[bool]string{true: "ok", false: "retained"}[change.applied],
		change.parseMode,
		change.fallbackReason,
	)
}

func sourceLineCount(source string) int {
	if source == "" {
		return 0
	}
	return strings.Count(source, "\n") + 1
}

func cacheStatus(hit bool) string {
	if hit {
		return "hit"
	}
	return "miss"
}
