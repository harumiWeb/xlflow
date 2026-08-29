package lspserver

import (
	"bytes"
	"context"
	"errors"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestPerformanceLoggingIsOptIn(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default(), Stderr: &output})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	measurement := s.startPerformance("textDocument/hover", intel.Document{URI: `file:///C:/work/日本語.bas`, Source: "Option Explicit\n"})
	if measurement != nil {
		t.Fatal("startPerformance returned a measurement while performance logging was disabled")
	}
	if strings.Contains(output.String(), "performance operation=") {
		t.Fatalf("performance output was emitted without opt-in: %s", output.String())
	}
}

func TestServerLogIdentifiesExecutableWithoutPerformanceLogging(t *testing.T) {
	var output bytes.Buffer
	_, cleanup, err := New(Options{
		RootDir: t.TempDir(), Config: config.Default(), Stderr: &output,
		Build: BuildInfo{Version: "dev", Commit: "abc123-dirty", Date: "2026-08-28"},
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	executable, err := os.Executable()
	if err != nil {
		t.Fatal(err)
	}
	for _, expected := range []string{strconv.Quote(executable), `version="dev"`, `commit="abc123-dirty"`, `build_date="2026-08-28"`} {
		if !strings.Contains(output.String(), expected) {
			t.Fatalf("server provenance missing %q: %s", expected, output.String())
		}
	}
}

func TestPerformanceLoggingIncludesStableDocumentFields(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	doc := intel.Document{
		URI:     `file:///C:/work/space%20and%20日本語.bas`,
		Path:    `C:\work\space and 日本語.bas`,
		Source:  "Option Explicit\nSub Main()\nEnd Sub\n",
		Version: 42,
	}
	s.startPerformance("textDocument/documentSymbol", doc).finish(3, nil)
	s.startPerformance("textDocument/hover", doc).finish(0, errors.New("boom"))

	logOutput := output.String()
	for _, expected := range []string{
		`performance operation="textDocument/documentSymbol"`,
		`uri="file:///C:/work/space%20and%20日本語.bas"`,
		`path="C:\\work\\space and 日本語.bas"`,
		`version=42`,
		`bytes=35`,
		`lines=4`,
		`result_count=3`,
		`outcome="ok"`,
		`performance operation="textDocument/hover"`,
		`outcome="error"`,
	} {
		if !strings.Contains(logOutput, expected) {
			t.Fatalf("performance log missing %q:\n%s", expected, logOutput)
		}
	}
}

func TestStartupTelemetryCorrelatesLifecycleReadinessAndInteractiveSuccess(t *testing.T) {
	fixture := makeLSPStartupBenchmarkFixture(t)
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        fixture.root,
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
		startup:        &startupContext{id: "attempt-test", started: time.Now()},
	})
	if err != nil {
		t.Fatal(err)
	}
	cleaned := false
	defer func() {
		if !cleaned {
			cleanup()
		}
	}()

	if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
		t.Fatal(err)
	}
	if err := s.initialized(nil, nil); err != nil {
		t.Fatal(err)
	}
	completionSource := fixture.sourceA
	if err := s.didOpen(nil, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(fixture.moduleAURI), Version: 1, Text: completionSource,
	}}); err != nil {
		t.Fatal(err)
	}
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}

	line := benchmarkSourceLine(completionSource, "    Call CrossFileTarget(1)")
	if _, err := s.hover(nil, &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(len("    Call "))},
	}}); err != nil {
		t.Fatal(err)
	}
	if _, err := s.definition(nil, &protocol.DefinitionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(len("    Call "))},
	}}); err != nil {
		t.Fatal(err)
	}
	completionLine := benchmarkSourceLine(completionSource, "    Worksheets(\"Input\").Ra")
	if _, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(completionLine), Character: 27},
	}}); err != nil {
		t.Fatal(err)
	}
	cleanup()
	cleaned = true

	logOutput := output.String()
	for _, event := range []string{
		"serverProcessStart", "serverConstructed", "initializeHandled", "initializedHandled",
		"didOpenHandled", "declarationIndexStarted", "declarationIndexReady", "semanticIndexStarted", "semanticIndexReady",
		"firstHoverHandled", "firstDefinitionHandled", "firstCompletionHandled",
	} {
		needle := `event="` + event + `"`
		if count := strings.Count(logOutput, needle); count != 1 {
			t.Fatalf("startup event %q count = %d, want one:\n%s", event, count, logOutput)
		}
	}
	for _, field := range []string{`startup_id="attempt-test"`, `elapsed_ms=`, `wall_time_unix_ns=`} {
		if !strings.Contains(logOutput, field) {
			t.Fatalf("startup telemetry missing %q:\n%s", field, logOutput)
		}
	}
}

func TestStartupTelemetryIsDisabledWithoutPerformanceLogging(t *testing.T) {
	var output bytes.Buffer
	_, cleanup, err := New(Options{
		RootDir: t.TempDir(), Config: config.Default(), Stderr: &output,
		startup: &startupContext{id: "attempt-disabled", started: time.Now()},
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	if strings.Contains(output.String(), `operation="lsp/startup"`) {
		t.Fatalf("startup telemetry emitted while performance logging was disabled: %s", output.String())
	}
}

func TestStartupContextReadsAnonymousEnvironmentIDWhenEnabled(t *testing.T) {
	t.Setenv(startupIDEnv, "attempt-from-client")
	if startup := startupContextFromEnvironment(false); startup != nil {
		t.Fatal("startup context was created while performance logging was disabled")
	}
	startup := startupContextFromEnvironment(true)
	if startup == nil || startup.id != "attempt-from-client" || startup.started.IsZero() {
		t.Fatalf("startup context = %+v, want enabled environment context", startup)
	}
	t.Setenv(startupIDEnv, " ")
	if startup := startupContextFromEnvironment(true); startup != nil {
		t.Fatalf("startup context = %+v, want nil for missing environment ID", startup)
	}
}

func TestDiagnosticsPerformanceLoggingIncludesGenerationAndDiscardStatus(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	s.startPerformance("diagnostics", intel.Document{
		URI:     "file:///work/Main.bas",
		Source:  "Option Explicit\n",
		Version: 9,
	}).finishDiagnostics(2, 12, true)

	logOutput := output.String()
	for _, expected := range []string{
		`operation="diagnostics"`,
		`version=9`,
		`generation=12`,
		`result_count=2`,
		`outcome="discarded"`,
		`discarded=true`,
	} {
		if !strings.Contains(logOutput, expected) {
			t.Fatalf("diagnostics performance log missing %q:\n%s", expected, logOutput)
		}
	}
}

func TestDocumentChangePerformanceLoggingIncludesParseModeAndFallback(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default(), Stderr: &output, PerformanceLog: true})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	s.logDocumentChangePerformance("file:///work/Main.bas", 4, documentChangeResult{
		document:       intel.Document{URI: "file:///work/Main.bas", Path: "/work/Main.bas", Source: "Sub Main()\nEnd Sub\n", Version: 4},
		applied:        true,
		parseMode:      "full_fallback",
		fallbackReason: "full_document_change",
	}, time.Now())
	for _, expected := range []string{
		`operation="textDocument/didChange/parse"`,
		`outcome="ok"`,
		`parse_mode="full_fallback"`,
		`fallback_reason="full_document_change"`,
	} {
		if !strings.Contains(output.String(), expected) {
			t.Fatalf("change performance log missing %q:\n%s", expected, output.String())
		}
	}
}

func TestDocumentsPreserveLSPVersion(t *testing.T) {
	docs := newDocuments(t.TempDir())
	uri := pathToFileURI(t.TempDir() + `/Main.bas`)
	doc, err := docs.open(uri, "Option Explicit\n", 7)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Version != 7 {
		t.Fatalf("open version = %d, want 7", doc.Version)
	}
	doc, err = docs.change(uri, "Option Explicit\nSub Main()\nEnd Sub\n", 8)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Version != 8 {
		t.Fatalf("change version = %d, want 8", doc.Version)
	}
}

func TestWorkspaceSymbolIndexPerformanceReportsInitialBuild(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	if _, err := s.cachedWorkspaceSymbols(nil, ""); err != nil {
		t.Fatal(err)
	}
	if err := s.analysis.waitReady(); err != nil {
		t.Fatal(err)
	}
	logOutput := output.String()
	if !strings.Contains(logOutput, `operation="workspaceSymbols/index/initial"`) ||
		!strings.Contains(logOutput, `file_count=0`) ||
		!strings.Contains(logOutput, `elapsed_ms=`) ||
		!strings.Contains(logOutput, `stage="workspaceDiscovery"`) ||
		!strings.Contains(logOutput, `outcome="counter_snapshot"`) {
		t.Fatalf("workspace index performance log missing initial build fields:\n%s", logOutput)
	}
	for _, name := range performanceCounterNames {
		if !strings.Contains(logOutput, `counter="`+name+`"`) {
			t.Fatalf("workspace index performance log missing counter %q:\n%s", name, logOutput)
		}
	}
}

func TestPerformanceRecorderRegistersDependencyCounters(t *testing.T) {
	var output bytes.Buffer
	recorder := newPerformanceRecorder(true, log.New(&output, "", 0))

	for _, counter := range []string{
		performanceCounterDependencyNodesUpdated,
		performanceCounterDependencyEdgesUpdated,
		performanceCounterProceduresRevisited,
	} {
		if got := recorder.counterTotal(counter); got != 0 {
			t.Fatalf("counter %q initialized to %d, want 0", counter, got)
		}
		recorder.addCounter(counter, 1, "workspace/project", performanceStageDependencyUpdate, "background", "")
		if got := recorder.counterTotal(counter); got != 1 {
			t.Fatalf("counter %q total = %d, want 1", counter, got)
		}
		if !strings.Contains(output.String(), `counter="`+counter+`" value=1 total=1`) {
			t.Fatalf("counter %q was not emitted:\n%s", counter, output.String())
		}
	}
}

func TestPerformanceRecorderRegistersInteractiveIndexCounters(t *testing.T) {
	var output bytes.Buffer
	recorder := newPerformanceRecorder(true, log.New(&output, "", 0))

	for _, counter := range []string{
		performanceCounterInteractiveIndexBuilds,
		performanceCounterInteractiveIndexHits,
		performanceCounterProcedureCatalogBuilds,
		performanceCounterProcedureCatalogReuses,
		performanceCounterInteractiveExactQueries,
		performanceCounterInteractivePrefixQueries,
		performanceCounterInteractiveQualifiedQueries,
		performanceCounterFullDocumentSymbolBuilds,
		performanceCounterInteractiveFullSymbolFallbacks,
	} {
		if got := recorder.counterTotal(counter); got != 0 {
			t.Fatalf("counter %q initialized to %d, want 0", counter, got)
		}
		recorder.addCounter(counter, 1, "textDocument/interactive", performanceStageDeclarationIndexing, "interactive", "")
		if got := recorder.counterTotal(counter); got != 1 {
			t.Fatalf("counter %q total = %d, want 1", counter, got)
		}
		if !strings.Contains(output.String(), `counter="`+counter+`" value=1 total=1`) {
			t.Fatalf("counter %q was not emitted:\n%s", counter, output.String())
		}
	}
}

func TestWorkspaceOverlayPerformanceReportsGenerationAndDiscard(t *testing.T) {
	var output bytes.Buffer
	s := &Server{opts: Options{PerformanceLog: true}, logger: log.New(&output, "", 0)}
	doc := intel.Document{URI: "file:///work/Main.bas", Path: "work/Main.bas", Source: "Sub Run()\nEnd Sub\n", Version: 7}
	s.logWorkspaceOverlayPerformance(doc, 12, time.Now(), nil, true)
	got := output.String()
	for _, field := range []string{`operation="workspaceSymbols/overlay"`, `generation=12`, `outcome="discarded"`, `discarded=true`} {
		if !strings.Contains(got, field) {
			t.Fatalf("overlay performance log missing %s:\n%s", field, got)
		}
	}
}

func TestDocumentSymbolCachePerformanceReportsMissThenHit(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	doc := intel.Document{
		URI:        `file:///C:/work/Main.bas`,
		Path:       `C:\work\Main.bas`,
		Source:     "Option Explicit\nSub Main()\nEnd Sub\n",
		ModuleKind: "standard",
		Version:    42,
	}
	doc = intel.NewAnalysisSnapshot(doc).Document()
	if _, err := s.analyzer.DocumentSymbols(doc); err != nil {
		t.Fatal(err)
	}
	if _, err := s.analyzer.DocumentSymbols(doc); err != nil {
		t.Fatal(err)
	}
	logOutput := output.String()
	for _, expected := range []string{
		`operation="documentSymbols/cache"`,
		`cache="miss"`,
		`cache="hit"`,
		`uri="file:///C:/work/Main.bas"`,
		`path="C:\\work\\Main.bas"`,
		`version=42`,
	} {
		if !strings.Contains(logOutput, expected) {
			t.Fatalf("document symbol cache log missing %q:\n%s", expected, logOutput)
		}
	}
}

func TestSemanticTokenCachePerformanceReportsMissHitAndGenerationTime(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	path := filepath.Join(t.TempDir(), "Main.bas")
	uri := pathToFileURI(path)
	if _, err := s.docs.open(uri, "Option Explicit\n", 3); err != nil {
		t.Fatal(err)
	}
	s.semanticTokenGenerator = func(context.Context, intel.Document, []intel.Document) ([]intel.SemanticToken, error) {
		return []intel.SemanticToken{{
			Range: intel.Range{Start: intel.Position{}, End: intel.Position{Character: 6}},
			Type:  intel.SemanticTokenKeyword,
		}}, nil
	}
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)}}
	if _, err := s.semanticTokensFull(nil, params); err != nil {
		t.Fatal(err)
	}
	if _, err := s.semanticTokensFull(nil, params); err != nil {
		t.Fatal(err)
	}

	logOutput := output.String()
	for _, expected := range []string{
		`operation="semanticTokens/cache"`,
		`cache="miss"`,
		`cache="hit"`,
		`operation="textDocument/semanticTokens/full"`,
		`elapsed_ms=`,
		`version=3`,
	} {
		if !strings.Contains(logOutput, expected) {
			t.Fatalf("semantic token performance log missing %q:\n%s", expected, logOutput)
		}
	}
}

func TestPerformanceLoggingIncludesDocumentResolutionFailures(t *testing.T) {
	var output bytes.Buffer
	s, cleanup, err := New(Options{
		RootDir:        t.TempDir(),
		Config:         config.Default(),
		Stderr:         &output,
		PerformanceLog: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	_, err = s.documentSymbol(nil, &protocol.DocumentSymbolParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: "https://example.invalid/Main.bas"},
	})
	if err == nil {
		t.Fatal("documentSymbol succeeded for an unsupported URI")
	}
	logOutput := output.String()
	for _, expected := range []string{
		`operation="textDocument/documentSymbol"`,
		`uri="https://example.invalid/Main.bas"`,
		`outcome="error"`,
	} {
		if !strings.Contains(logOutput, expected) {
			t.Fatalf("resolution failure log missing %q:\n%s", expected, logOutput)
		}
	}
}

func TestSourceLineCount(t *testing.T) {
	tests := []struct {
		source string
		want   int
	}{
		{"", 0},
		{"Option Explicit", 1},
		{"Option Explicit\n", 2},
		{"a\nb\nc", 3},
	}
	for _, test := range tests {
		if got := sourceLineCount(test.source); got != test.want {
			t.Errorf("sourceLineCount(%q) = %d, want %d", test.source, got, test.want)
		}
	}
}

func TestLSPPreparationTelemetryReportsStagesAndCounters(t *testing.T) {
	var output bytes.Buffer
	s := &Server{performance: newPerformanceRecorder(true, log.New(&output, "", 0))}
	project := projectTestSnapshot(projectTestProcedure(filepath.Join(t.TempDir(), "Main.bas"), "Main.Run", "", "", 1, "Run"))
	project.Revision = 19

	ctx := context.Background()
	if _, _, _, _, err := s.projectResolution(ctx, project, true); err != nil {
		t.Fatal(err)
	}
	if _, _, _, _, err := s.projectResolution(ctx, project, true); err != nil {
		t.Fatal(err)
	}
	s.projectEffectSummaryWithResolution(ctx, project, nil, true)
	s.projectEffectSummaryWithResolution(ctx, project, nil, true)
	s.projectConstants(project, true, nil)
	s.projectConstants(project, true, nil)

	for _, expected := range []string{
		`stage="projectResolver"`,
		`stage="projectResolutionView"`,
		`stage="projectResolutionMaterialization"`,
		`stage="projectEffectSummary"`,
		`stage="projectConstants"`,
		`counter="resolution_resolver_builds" value=1`,
		`counter="resolution_view_builds" value=1`,
		`counter="canonical_resolver_builds" value=1`,
		`counter="procedure_resolver_views" value=1`,
		`counter="full_resolver_views" value=1`,
		`counter="resolution_overlay_builds" value=2`,
	} {
		if !strings.Contains(output.String(), expected) {
			t.Fatalf("preparation telemetry missing %q:\n%s", expected, output.String())
		}
	}
	if got := s.performance.counterTotal(performanceCounterResolutionResolverBuilds); got != 1 {
		t.Fatalf("resolver build counter = %d, want 1", got)
	}
	if got := s.performance.counterTotal(performanceCounterResolutionMaterializations); got != 0 {
		t.Fatalf("resolution materializations = %d, want 0", got)
	}
	if got := s.performance.counterTotal(performanceCounterCanonicalResolverBuilds); got != 1 {
		t.Fatalf("canonical resolver builds = %d, want 1", got)
	}
	if got := s.performance.counterTotal(performanceCounterProcedureResolverViews); got != 1 {
		t.Fatalf("procedure resolver views = %d, want 1", got)
	}
	if got := s.performance.counterTotal(performanceCounterFullResolverViews); got != 1 {
		t.Fatalf("full resolver views = %d, want 1", got)
	}
	if got := s.performance.counterTotal(performanceCounterResolutionOverlayBuilds); got != 2 {
		t.Fatalf("resolution overlay builds = %d, want 2", got)
	}
}

func TestAnalysisPermitPreCanceledContextDoesNotWait(t *testing.T) {
	var output bytes.Buffer
	s := &Server{
		performance:       newPerformanceRecorder(true, log.New(&output, "", 0)),
		analysisScheduler: newAnalysisScheduler(1),
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if release, _, err := s.acquireAnalysisPermit(ctx, analysisWorkBackground); release != nil || !errors.Is(err, context.Canceled) {
		t.Fatalf("acquire canceled = (release=%v, err=%v)", release != nil, err)
	}
	if got := s.performance.counterTotal(performanceCounterBackgroundPermitWaits); got != 0 {
		t.Fatalf("background permit waits = %d, want 0 for a pre-canceled context", got)
	}
	if strings.Contains(output.String(), `stage="permitWait"`) {
		t.Fatalf("pre-canceled permit emitted a wait record: %s", output.String())
	}
}

func TestAnalysisPermitFastWaitDoesNotIncrementInteractiveOrBackgroundCounters(t *testing.T) {
	var output bytes.Buffer
	s := &Server{
		performance:       newPerformanceRecorder(true, log.New(&output, "", 0)),
		analysisScheduler: newAnalysisScheduler(1),
	}
	holder, _, err := s.acquireAnalysisPermit(context.Background(), analysisWorkFast)
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan error, 1)
	go func() {
		_, _, acquireErr := s.acquireAnalysisPermit(ctx, analysisWorkFast)
		done <- acquireErr
	}()
	waitForAnalysisWaiter(t, s.analysisScheduler, analysisWorkFast)
	cancel()
	select {
	case err := <-done:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("Fast wait error = %v, want context.Canceled", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("Fast waiter did not observe cancellation")
	}
	holder()
	if got := s.performance.counterTotal(performanceCounterInteractivePermitWaits); got != 0 {
		t.Fatalf("interactive permit waits = %d, want 0 for Fast work", got)
	}
	if got := s.performance.counterTotal(performanceCounterBackgroundPermitWaits); got != 0 {
		t.Fatalf("background permit waits = %d, want 0 for Fast work", got)
	}
	if !strings.Contains(output.String(), `class="fast"`) || !strings.Contains(output.String(), `analysis_permit_wait_ms=`) {
		t.Fatalf("Fast permit telemetry missing wait record:\n%s", output.String())
	}
}

func TestAnalysisPermitInteractivePriority(t *testing.T) {
	var output bytes.Buffer
	s := &Server{
		performance:       newPerformanceRecorder(true, log.New(&output, "", 0)),
		analysisScheduler: newAnalysisScheduler(2),
	}
	backgroundHolderRelease, _, err := s.acquireAnalysisPermit(context.Background(), analysisWorkBackground)
	if err != nil {
		t.Fatal(err)
	}
	fastHolderRelease, _, err := s.acquireAnalysisPermit(context.Background(), analysisWorkFast)
	if err != nil {
		t.Fatal(err)
	}

	backgroundDone := make(chan func(), 1)
	go func() {
		release, _, acquireErr := s.acquireAnalysisPermit(context.Background(), analysisWorkBackground)
		if acquireErr != nil {
			return
		}
		backgroundDone <- release
	}()
	waitForAnalysisWaiter(t, s.analysisScheduler, analysisWorkBackground)
	select {
	case release := <-backgroundDone:
		release()
		t.Fatal("background waiter acquired the held permit")
	default:
	}

	interactiveDone := make(chan func(), 1)
	go func() {
		release, _, acquireErr := s.acquireAnalysisPermit(context.Background(), analysisWorkInteractive)
		if acquireErr != nil {
			return
		}
		interactiveDone <- release
	}()
	waitForAnalysisWaiter(t, s.analysisScheduler, analysisWorkInteractive)
	if s.analysisScheduler.waiterCount(analysisWorkInteractive) == 0 {
		t.Fatal("interactive waiter did not register")
	}

	fastHolderRelease()
	var interactiveRelease func()
	select {
	case interactiveRelease = <-interactiveDone:
	case <-time.After(time.Second):
		t.Fatal("interactive waiter did not acquire the released permit")
	}
	select {
	case release := <-backgroundDone:
		release()
		t.Fatal("background waiter bypassed the interactive waiter")
	default:
	}
	backgroundHolderRelease()
	interactiveRelease()
	select {
	case release := <-backgroundDone:
		release()
	case <-time.After(time.Second):
		t.Fatal("background waiter did not resume after interactive release")
	}
	for _, expected := range []string{
		`interactive_wait_ms=`,
		`background_wait_ms=`,
		`analysis_permit_wait_ms=`,
		`current_workers=`,
		`max_active_workers=`,
	} {
		if !strings.Contains(output.String(), expected) {
			t.Fatalf("analysis scheduler telemetry missing %q:\n%s", expected, output.String())
		}
	}
}

func waitForAnalysisWaiter(t *testing.T, scheduler *analysisScheduler, class analysisWorkClass) {
	t.Helper()
	deadline := time.NewTimer(time.Second)
	defer deadline.Stop()
	ticker := time.NewTicker(time.Millisecond)
	defer ticker.Stop()
	for {
		if scheduler.waiterCount(class) > 0 {
			return
		}
		select {
		case <-ticker.C:
		case <-deadline.C:
			t.Fatalf("%s waiter did not register", class)
		}
	}
}
