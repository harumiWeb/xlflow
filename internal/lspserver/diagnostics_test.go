package lspserver

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"time"
	"unicode/utf16"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

func TestDiagnosticsChangesCoalesceToLatestDebouncedGeneration(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	versions := make(chan int32, 4)
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		versions <- doc.Version
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	wantVersion(t, versions, 1)

	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 3)
	created := timers.snapshot()
	if len(created) != 4 {
		t.Fatalf("timers = %d, want 4 (Fast and Full for each change)", len(created))
	}
	created[0].Fire()
	wantNoVersion(t, versions)
	created[2].Fire()
	wantVersion(t, versions, 3)
}

func TestDidChangeSchedulesFastThenIdleFullDiagnostics(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	runs := make(chan int32, 4)
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		runs <- doc.Version
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	wantVersion(t, runs, 1)

	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	created := timers.snapshot()
	if len(created) != 2 {
		t.Fatalf("timers = %d, want Fast and Full", len(created))
	}
	if created[0].delay != diagnosticsDebounce || created[1].delay != diagnosticsFullIdleDelay {
		t.Fatalf("timer delays = %v, %v", created[0].delay, created[1].delay)
	}
	created[0].Fire()
	wantVersion(t, runs, 2)
	created[1].Fire()
	wantVersion(t, runs, 2)
}

func TestFastDiagnosticsAreCompletelyReplacedByFullDiagnostics(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	s.diagnosticsRequest = func(_ context.Context, request intel.DiagnosticRequest) intel.DiagnosticResult {
		return intel.DiagnosticResult{Diagnostics: []intel.Diagnostic{{Code: "FAST", Severity: "warning", Message: fmt.Sprintf("fast %d", request.Document.Version)}}}
	}
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		return []intel.Diagnostic{{Code: "FULL", Severity: "warning", Message: fmt.Sprintf("full %d", doc.Version)}}
	}
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	notifications.waitForCount(t, 1)
	notifications.clear()

	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	created := timers.snapshot()
	created[0].Fire()
	fast := notifications.waitForCount(t, 1)
	if len(fast[0].Diagnostics) != 1 || fast[0].Diagnostics[0].Message != "fast 2" {
		t.Fatalf("Fast publication = %+v", fast)
	}
	waitDiagnosticsIdle(t, s, uri)

	created[1].Fire()
	publications := notifications.waitForCount(t, 2)
	full := publications[1]
	if len(full.Diagnostics) != 1 || full.Diagnostics[0].Message != "full 2" {
		t.Fatalf("Full publication did not replace Fast exactly: %+v", publications)
	}
}

func TestEditingCancelsRunningFullBeforeLatestFastPublishes(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	fullStarted := make(chan struct{})
	fullCanceled := make(chan struct{})
	s.diagnosticsRequest = func(_ context.Context, request intel.DiagnosticRequest) intel.DiagnosticResult {
		return intel.DiagnosticResult{Diagnostics: []intel.Diagnostic{{Code: "FAST", Severity: "warning", Message: fmt.Sprintf("fast %d", request.Document.Version)}}}
	}
	s.diagnostics = func(runCtx context.Context, doc intel.Document) []intel.Diagnostic {
		if doc.Version == 2 {
			close(fullStarted)
			<-runCtx.Done()
			close(fullCanceled)
		}
		return []intel.Diagnostic{{Code: "FULL", Severity: "warning", Message: fmt.Sprintf("full %d", doc.Version)}}
	}
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	notifications.waitForCount(t, 1)
	notifications.clear()

	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	created := timers.snapshot()
	created[0].Fire()
	notifications.waitForCount(t, 1)
	waitDiagnosticsIdle(t, s, uri)
	created[1].Fire()
	waitClosed(t, fullStarted, "Full diagnostics")

	changeDiagnosticsTestDocument(t, s, ctx, uri, 3)
	created = timers.snapshot()
	created[2].Fire()
	waitClosed(t, fullCanceled, "obsolete Full cancellation")
	publications := notifications.waitForCount(t, 2)
	if len(publications) != 2 || len(publications[1].Diagnostics) != 1 || publications[1].Diagnostics[0].Message != "fast 3" {
		t.Fatalf("publications after Full cancellation = %+v, want Fast version 3 without obsolete Full", publications)
	}
}

func TestDidOpenReturnsBeforeBackgroundAnalysisCompletes(t *testing.T) {
	s, _, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan struct{})
	release := make(chan struct{})
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic {
		close(started)
		<-release
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, "Main.bas"))
	returned := make(chan error, 1)
	go func() {
		returned <- s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
			URI: protocol.DocumentUri(uri), Version: 1, Text: diagnosticSource(1),
		}})
	}()
	select {
	case err := <-returned:
		if err != nil {
			t.Fatal(err)
		}
	case <-time.After(time.Second):
		t.Fatal("didOpen waited for background analysis")
	}
	waitClosed(t, started, "background diagnostics")
	close(release)
}

func TestLSPDiagnosticsIncludeVBA215FromSharedRealtimeAnalysis(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	diagnostics := s.diagnostics(context.Background(), intel.Document{
		Path: filepath.Join(root, "Main.bas"),
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim rng As Range
    rng.Find _
        What:="missing"
End Sub
`,
	})
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == "VBA215" {
			if diagnostic.Range.Start.Line != 4 || diagnostic.Range.End.Line != 5 || diagnostic.Message != "Range.Find omits stateful argument(s): LookIn, LookAt, SearchOrder, MatchByte." {
				t.Fatalf("VBA215 diagnostic = %+v", diagnostic)
			}
			return
		}
	}
	t.Fatalf("VBA215 diagnostic missing: %+v", diagnostics)
}

func TestChangedProcedureNamesDetectsByRefSignatureEdits(t *testing.T) {
	before := map[string]procedureSignature{
		"helper.takevalue.sub": {name: "TakeValue", fingerprint: "sub||public|value:long:false:byref:false:false"},
	}
	after := map[string]procedureSignature{
		"helper.takevalue.sub": {name: "TakeValue", fingerprint: "sub||public|value:string:false:byref:false:false"},
	}
	changed := changedProcedureNames(before, after)
	if len(changed) != 1 || changed[0] != "TakeValue" {
		t.Fatalf("changed procedure names = %+v, want TakeValue", changed)
	}
}

func TestFirstOpenProcedureRemovalSchedulesDependentCallerDiagnostics(t *testing.T) {
	root := t.TempDir()
	modules := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(modules, 0o755); err != nil {
		t.Fatal(err)
	}
	callerPath := filepath.Join(modules, "Caller.bas")
	callerSource := "Attribute VB_Name = \"Caller\"\nOption Explicit\nPublic Sub Run()\n  Dim text As String\n  TakeLong (text)\nEnd Sub\n"
	calleePath := filepath.Join(modules, "Callee.bas")
	diskCallee := "Attribute VB_Name = \"Callee\"\nOption Explicit\nPublic Sub TakeLong(ByRef value As Long)\nEnd Sub\n"
	unsavedCallee := "Attribute VB_Name = \"Callee\"\nOption Explicit\nPublic Sub TakeText(ByRef value As String)\nEnd Sub\n"
	if err := os.WriteFile(callerPath, []byte(callerSource), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(calleePath, []byte(diskCallee), 0o644); err != nil {
		t.Fatal(err)
	}
	cfg := config.Default()
	cfg.Analyze.DetectByRefArgumentMismatch = true
	s, cleanup, err := New(Options{RootDir: root, Config: cfg})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	runs := make(chan string, 16)
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		runs <- filepath.Base(doc.Path)
		return nil
	}
	ctx := diagnosticTestContext(nil)
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(pathToFileURI(callerPath)), Version: 1, Text: callerSource,
	}}); err != nil {
		t.Fatal(err)
	}
	wantDiagnosticRun(t, runs, "Caller.bas")
	drainDiagnosticRuns(runs)

	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(pathToFileURI(calleePath)), Version: 1, Text: unsavedCallee,
	}}); err != nil {
		t.Fatal(err)
	}
	wantDiagnosticRun(t, runs, "Callee.bas")
	wantDiagnosticRun(t, runs, "Caller.bas")
	drainDiagnosticRuns(runs)
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(pathToFileURI(calleePath))},
	}); err != nil {
		t.Fatal(err)
	}
	wantDiagnosticRun(t, runs, "Caller.bas")
}

func wantDiagnosticRun(t *testing.T, runs <-chan string, want string) {
	t.Helper()
	deadline := time.After(5 * time.Second)
	for {
		select {
		case got := <-runs:
			if got == want {
				return
			}
		case <-deadline:
			t.Fatalf("timed out waiting for diagnostics on %s", want)
		}
	}
}

func drainDiagnosticRuns(runs <-chan string) {
	for {
		select {
		case <-runs:
		default:
			return
		}
	}
}

func TestLSPDiagnosticsIncludeWorksheetRootFindings(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	diagnostics := s.diagnostics(context.Background(), intel.Document{
		Path: filepath.Join(root, "Main.bas"),
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub Run()
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRow As Long
    Set inputSheet = ThisWorkbook.Worksheets("Input")
    Set outputSheet = ThisWorkbook.Worksheets("Output")
    lastRow = inputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row
    lastRow = Cells(Rows.Count, 1).End(xlDown).Row
End Sub
`,
	})
	var mismatch, unstable int
	for _, diagnostic := range diagnostics {
		switch diagnostic.Code {
		case "VBA216":
			if diagnostic.Severity != "error" || diagnostic.Range.Start.Line != 8 {
				t.Fatalf("VBA216 diagnostic = %+v", diagnostic)
			}
			mismatch++
		case "VBA217":
			if diagnostic.Severity != "warning" || diagnostic.Range.Start.Line != 9 {
				t.Fatalf("VBA217 diagnostic = %+v", diagnostic)
			}
			unstable++
		}
	}
	if mismatch != 1 || unstable != 2 {
		t.Fatalf("worksheet diagnostics: VBA216=%d, VBA217=%d; all=%+v", mismatch, unstable, diagnostics)
	}
}

func TestLSPDiagnosticsPreserveVBA215UTF16StartRange(t *testing.T) {
	root := t.TempDir()
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()

	line := `    Debug.Print "😀日本": Set found = rng.Find("missing")`
	diagnostics := s.diagnostics(context.Background(), intel.Document{
		Path:   filepath.Join(root, "Main.bas"),
		Source: "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n    Dim rng As Range\n    Dim found As Range\n" + line + "\nEnd Sub\n",
	})
	wantStart := len(utf16.Encode([]rune(line[:strings.Index(line, "rng.Find")])))
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == "VBA215" {
			if diagnostic.Range.Start.Line != 5 || diagnostic.Range.Start.Character != wantStart {
				t.Fatalf("VBA215 UTF-16 range = %+v, want line 5 character %d", diagnostic.Range, wantStart)
			}
			return
		}
	}
	t.Fatalf("VBA215 diagnostic missing: %+v", diagnostics)
}

func TestDiagnosticsChangeDuringDidOpenAnalysisBecomesLatestPendingGeneration(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan int32, 2)
	canceled := make(chan struct{})
	release := make(chan struct{})
	var releaseOnce sync.Once
	releaseWorker := func() { releaseOnce.Do(func() { close(release) }) }
	defer releaseWorker()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	s.diagnostics = func(runCtx context.Context, doc intel.Document) []intel.Diagnostic {
		started <- doc.Version
		if doc.Version == 1 {
			<-runCtx.Done()
			close(canceled)
			<-release
		}
		return diagnosticVersionResult(runCtx, doc)
	}
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, "Main.bas"))
	openDone := make(chan struct{})
	go func() {
		_ = s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
			URI: protocol.DocumentUri(uri), Version: 1, Text: diagnosticSource(1),
		}})
		close(openDone)
	}()
	wantVersion(t, started, 1)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	waitClosed(t, canceled, "didOpen diagnostics cancellation")
	wantNoVersion(t, started)
	releaseWorker()
	waitClosed(t, openDone, "didOpen")
	wantVersion(t, started, 2)

	params := notifications.waitForCount(t, 1)
	if len(params) != 1 || len(params[0].Diagnostics) != 1 || params[0].Diagnostics[0].Message != "version 2" {
		t.Fatalf("published diagnostics = %+v, want only version 2", params)
	}
}

func TestDiagnosticsKeepsOneWorkerAndRunsOnlyLatestReadyGeneration(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan int32, 4)
	release := make(chan struct{})
	var active atomic.Int64
	var maximum atomic.Int64
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		current := active.Add(1)
		defer active.Add(-1)
		for {
			observed := maximum.Load()
			if current <= observed || maximum.CompareAndSwap(observed, current) {
				break
			}
		}
		started <- doc.Version
		if doc.Version == 2 {
			<-release
		}
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	wantVersion(t, started, 1)

	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	wantVersion(t, started, 2)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 3)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 4)
	timers.snapshot()[4].Fire()
	wantNoVersion(t, started)
	close(release)
	wantVersion(t, started, 4)
	if maximum.Load() != 1 {
		t.Fatalf("maximum concurrent workers = %d, want 1", maximum.Load())
	}
}

func TestDiagnosticsDiscardsGenerationChangedImmediatelyBeforePublish(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	s.diagnostics = diagnosticVersionResult
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	notifications.clear()

	hookStarted := make(chan struct{})
	releaseHook := make(chan struct{})
	var once sync.Once
	s.beforeDiagnosticsPublish = func() {
		once.Do(func() {
			close(hookStarted)
			<-releaseHook
		})
	}
	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	waitClosed(t, hookStarted, "publish hook")
	changeDiagnosticsTestDocument(t, s, ctx, uri, 3)
	close(releaseHook)
	timers.snapshot()[2].Fire()

	params := notifications.waitForCount(t, 1)
	if len(params) != 1 || len(params[0].Diagnostics) != 1 || params[0].Diagnostics[0].Message != "version 3" {
		t.Fatalf("published diagnostics = %+v, want only version 3", params)
	}
}

func TestDiagnosticsCloseMakesEmptyPublishFinal(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	notifications := &diagnosticNotificationRecorder{}
	ctx := diagnosticTestContext(notifications)
	release := make(chan struct{})
	started := make(chan struct{})
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		if doc.Version == 2 {
			close(started)
			<-release
		}
		return diagnosticVersionResult(context.Background(), doc)
	}
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	notifications.waitForCount(t, 1)
	notifications.clear()
	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	waitClosed(t, started, "diagnostics worker")
	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	}); err != nil {
		t.Fatal(err)
	}
	close(release)
	s.stopDiagnostics()

	params := notifications.snapshot()
	if len(params) != 1 || len(params[0].Diagnostics) != 0 {
		t.Fatalf("notifications after close = %+v, want one final empty publish", params)
	}
}

func TestDiagnosticsReopenWaitsForCanceledWorkerOnSameURI(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan int32, 4)
	release := make(chan struct{})
	var active atomic.Int64
	var maximum atomic.Int64
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		current := active.Add(1)
		defer active.Add(-1)
		for {
			observed := maximum.Load()
			if current <= observed || maximum.CompareAndSwap(observed, current) {
				break
			}
		}
		started <- doc.Version
		if doc.Version == 2 {
			<-release
		}
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	wantVersion(t, started, 1)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	wantVersion(t, started, 2)

	if err := s.didClose(ctx, &protocol.DidCloseTextDocumentParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
	}); err != nil {
		t.Fatal(err)
	}
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), Version: 3, Text: diagnosticSource(3),
	}}); err != nil {
		t.Fatal(err)
	}
	wantNoVersion(t, started)
	close(release)
	wantVersion(t, started, 3)
	if maximum.Load() != 1 {
		t.Fatalf("maximum concurrent workers across reopen = %d, want 1", maximum.Load())
	}
}

func TestDocumentLifecycleSerializesCloseAndReopen(t *testing.T) {
	s, _, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	s.diagnostics = func(context.Context, intel.Document) []intel.Diagnostic { return nil }
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)

	notifyStarted := make(chan struct{})
	releaseNotify := make(chan struct{})
	closeCtx := &glsp.Context{Notify: func(_ string, _ any) {
		close(notifyStarted)
		<-releaseNotify
	}}
	closeDone := make(chan struct{})
	go func() {
		_ = s.didClose(closeCtx, &protocol.DidCloseTextDocumentParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
		})
		close(closeDone)
	}()
	waitClosed(t, notifyStarted, "close notification")

	openDone := make(chan struct{})
	go func() {
		_ = s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
			URI: protocol.DocumentUri(uri), Version: 3, Text: diagnosticSource(3),
		}})
		close(openDone)
	}()
	select {
	case <-openDone:
		t.Fatal("reopen completed while close still owned the URI lifecycle")
	default:
	}
	close(releaseNotify)
	waitClosed(t, closeDone, "close")
	waitClosed(t, openDone, "reopen")
	doc, err := s.docs.getOrRead(uri)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Version != 3 || doc.Source != diagnosticSource(3) {
		t.Fatalf("reopened document = %+v, want version 3 buffer", doc)
	}
}

func TestDiagnosticsSlowDocumentDoesNotBlockAnotherDocument(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan string, 4)
	releaseA := make(chan struct{})
	s.diagnostics = func(_ context.Context, doc intel.Document) []intel.Diagnostic {
		started <- doc.URI
		if doc.Version == 2 && filepath.Base(doc.Path) == "A.bas" {
			<-releaseA
		}
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uriA := openDiagnosticsTestDocumentAt(t, s, ctx, "A.bas", 1)
	wantURI(t, started, uriA)
	uriB := openDiagnosticsTestDocumentAt(t, s, ctx, "B.bas", 1)
	wantURI(t, started, uriB)

	changeDiagnosticsTestDocument(t, s, ctx, uriA, 2)
	timers.snapshot()[0].Fire()
	wantURI(t, started, uriA)
	changeDiagnosticsTestDocument(t, s, ctx, uriB, 2)
	timers.snapshot()[2].Fire()
	wantURI(t, started, uriB)
	close(releaseA)
}

func TestDiagnosticsShutdownStopsTimersCancelsAndWaitsForWorkers(t *testing.T) {
	s, timers, cleanup := newDiagnosticsTestServer(t)
	defer cleanup()
	started := make(chan struct{})
	canceled := make(chan struct{})
	release := make(chan struct{})
	s.diagnostics = func(ctx context.Context, doc intel.Document) []intel.Diagnostic {
		if doc.Version == 2 {
			close(started)
			<-ctx.Done()
			close(canceled)
			<-release
		}
		return nil
	}
	ctx := diagnosticTestContext(nil)
	uri := openDiagnosticsTestDocument(t, s, ctx, 1)
	changeDiagnosticsTestDocument(t, s, ctx, uri, 2)
	timers.snapshot()[0].Fire()
	waitClosed(t, started, "diagnostics worker")
	changeDiagnosticsTestDocument(t, s, ctx, uri, 3)
	pending := timers.snapshot()[2]

	shutdownDone := make(chan struct{})
	go func() {
		_ = s.shutdown(nil)
		close(shutdownDone)
	}()
	waitClosed(t, canceled, "diagnostics cancellation")
	waitClosed(t, pending.stoppedCh, "pending timer stop")
	select {
	case <-shutdownDone:
		t.Fatal("shutdown returned before the active worker exited")
	default:
	}
	close(release)
	waitClosed(t, shutdownDone, "shutdown")
}

type fakeDiagnosticTimers struct {
	mu     sync.Mutex
	timers []*fakeDiagnosticTimer
}

func (f *fakeDiagnosticTimers) AfterFunc(delay time.Duration, callback func()) diagnosticTimer {
	timer := &fakeDiagnosticTimer{delay: delay, callback: callback, stoppedCh: make(chan struct{})}
	f.mu.Lock()
	f.timers = append(f.timers, timer)
	f.mu.Unlock()
	return timer
}

func (f *fakeDiagnosticTimers) snapshot() []*fakeDiagnosticTimer {
	f.mu.Lock()
	defer f.mu.Unlock()
	return append([]*fakeDiagnosticTimer(nil), f.timers...)
}

type fakeDiagnosticTimer struct {
	mu        sync.Mutex
	delay     time.Duration
	callback  func()
	stopped   bool
	stoppedCh chan struct{}
}

func (f *fakeDiagnosticTimer) Stop() bool {
	f.mu.Lock()
	defer f.mu.Unlock()
	wasActive := !f.stopped
	f.stopped = true
	if wasActive {
		close(f.stoppedCh)
	}
	return wasActive
}

func (f *fakeDiagnosticTimer) Fire() {
	f.mu.Lock()
	callback := f.callback
	f.mu.Unlock()
	callback()
}

type diagnosticNotificationRecorder struct {
	mu      sync.Mutex
	params  []protocol.PublishDiagnosticsParams
	changed chan struct{}
}

func (r *diagnosticNotificationRecorder) add(params protocol.PublishDiagnosticsParams) {
	r.mu.Lock()
	r.params = append(r.params, params)
	if r.changed != nil {
		close(r.changed)
		r.changed = nil
	}
	r.mu.Unlock()
}

func (r *diagnosticNotificationRecorder) clear() {
	r.mu.Lock()
	r.params = nil
	r.mu.Unlock()
}

func (r *diagnosticNotificationRecorder) snapshot() []protocol.PublishDiagnosticsParams {
	r.mu.Lock()
	defer r.mu.Unlock()
	return append([]protocol.PublishDiagnosticsParams(nil), r.params...)
}

func (r *diagnosticNotificationRecorder) waitForCount(t *testing.T, count int) []protocol.PublishDiagnosticsParams {
	t.Helper()
	deadline := time.NewTimer(2 * time.Second)
	defer deadline.Stop()
	for {
		r.mu.Lock()
		if len(r.params) >= count {
			params := append([]protocol.PublishDiagnosticsParams(nil), r.params...)
			r.mu.Unlock()
			return params
		}
		if r.changed == nil {
			r.changed = make(chan struct{})
		}
		changed := r.changed
		r.mu.Unlock()
		select {
		case <-changed:
		case <-deadline.C:
			t.Fatalf("diagnostic notifications = %+v, want at least %d", r.snapshot(), count)
		}
	}
}

func newDiagnosticsTestServer(t *testing.T) (*Server, *fakeDiagnosticTimers, func()) {
	t.Helper()
	s, cleanup, err := New(Options{RootDir: t.TempDir(), Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	timers := &fakeDiagnosticTimers{}
	s.diagnosticsAfterFunc = timers.AfterFunc
	s.diagnosticsOpenDelay = 0
	// Scheduler tests replace the legacy diagnostic seam after construction.
	// Route both phases through that dynamic seam unless a test explicitly
	// exercises the real Fast/Full analyzer.
	s.diagnosticsRequest = func(ctx context.Context, request intel.DiagnosticRequest) intel.DiagnosticResult {
		return intel.DiagnosticResult{Diagnostics: s.diagnostics(ctx, request.Document)}
	}
	return s, timers, cleanup
}

func diagnosticTestContext(recorder *diagnosticNotificationRecorder) *glsp.Context {
	return &glsp.Context{Notify: func(_ string, params any) {
		if recorder == nil {
			return
		}
		publish, ok := params.(protocol.PublishDiagnosticsParams)
		if ok {
			recorder.add(publish)
		}
	}}
}

func openDiagnosticsTestDocument(t *testing.T, s *Server, ctx *glsp.Context, version int32) string {
	t.Helper()
	return openDiagnosticsTestDocumentAt(t, s, ctx, "Main.bas", version)
}

func openDiagnosticsTestDocumentAt(t *testing.T, s *Server, ctx *glsp.Context, name string, version int32) string {
	t.Helper()
	uri := pathToFileURI(filepath.Join(s.opts.RootDir, name))
	if err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(uri), Version: version, Text: diagnosticSource(version),
	}}); err != nil {
		t.Fatal(err)
	}
	return uri
}

func changeDiagnosticsTestDocument(t *testing.T, s *Server, ctx *glsp.Context, uri string, version int32) {
	t.Helper()
	if err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
		TextDocument: protocol.VersionedTextDocumentIdentifier{
			TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(uri)},
			Version:                version,
		},
		ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: diagnosticSource(version)}},
	}); err != nil {
		t.Fatal(err)
	}
}

func diagnosticSource(version int32) string {
	return fmt.Sprintf("Option Explicit\nSub Version%d()\nEnd Sub\n", version)
}

func diagnosticVersionResult(_ context.Context, doc intel.Document) []intel.Diagnostic {
	return []intel.Diagnostic{{Code: "TEST", Severity: "warning", Source: "xlflow", Message: fmt.Sprintf("version %d", doc.Version)}}
}

func wantVersion(t *testing.T, versions <-chan int32, want int32) {
	t.Helper()
	select {
	case got := <-versions:
		if got != want {
			t.Fatalf("diagnostics version = %d, want %d", got, want)
		}
	case <-time.After(2 * time.Second):
		t.Fatalf("timed out waiting for diagnostics version %d", want)
	}
}

func wantNoVersion(t *testing.T, versions <-chan int32) {
	t.Helper()
	select {
	case got := <-versions:
		t.Fatalf("unexpected diagnostics version %d", got)
	default:
	}
}

func wantURI(t *testing.T, uris <-chan string, want string) {
	t.Helper()
	select {
	case got := <-uris:
		if got != want {
			t.Fatalf("diagnostics URI = %q, want %q", got, want)
		}
	case <-time.After(2 * time.Second):
		t.Fatalf("timed out waiting for diagnostics URI %q", want)
	}
}

func waitClosed(t *testing.T, ch <-chan struct{}, label string) {
	t.Helper()
	select {
	case <-ch:
	case <-time.After(2 * time.Second):
		t.Fatalf("timed out waiting for %s", label)
	}
}

func waitDiagnosticsIdle(t *testing.T, s *Server, uri string) {
	t.Helper()
	deadline := time.Now().Add(2 * time.Second)
	for time.Now().Before(deadline) {
		s.diagMu.Lock()
		state := s.diagStates[uri]
		s.diagMu.Unlock()
		if state != nil {
			state.mu.Lock()
			running := state.running
			state.mu.Unlock()
			if !running {
				return
			}
		}
		time.Sleep(5 * time.Millisecond)
	}
	t.Fatalf("diagnostics for %s did not become idle", uri)
}
