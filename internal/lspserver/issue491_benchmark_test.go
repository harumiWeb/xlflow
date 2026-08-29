package lspserver

import (
	"context"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"sync/atomic"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const (
	issue491ProcedureCount = 1200
	issue491CallsPerProc   = 4
	// The first-publication benchmark includes the initial Fast pass over the
	// complete synthetic module. Keep a generous guard so a contended host
	// cannot leave the benchmark process waiting forever.
	issue491FirstPublicationWait = 5 * time.Minute
	// Building the 1,200-procedure catalog can exceed five seconds on a
	// contended Linux CI runner. The behavior under test is prompt cancellation
	// after the checkpoint, not a fixed upper bound on catalog construction.
	issue491CheckpointWait = 30 * time.Second
)

// issue491LargeClassSource is deliberately generated rather than checked in:
// it keeps the repository independent of the third-party ROneCOne source while
// retaining the procedure, call-site, and line-count scale that exposed #491.
func issue491LargeClassSource() string {
	var out strings.Builder
	out.Grow(900_000)
	out.WriteString("VERSION 1.0 CLASS\n")
	out.WriteString("BEGIN\n  MultiUse = -1\nEND\n")
	out.WriteString("Attribute VB_Name = \"Issue491LargeClass\"\n")
	out.WriteString("Option Explicit\n\n")
	for procedure := 0; procedure < issue491ProcedureCount; procedure++ {
		fmt.Fprintf(&out, "Public Sub MassiveProcedure%04d(ByVal value As Long)\n", procedure)
		out.WriteString("    Dim worksheetRef As Worksheet\n")
		out.WriteString("    Dim cellRef As Range\n")
		out.WriteString("    Set worksheetRef = ThisWorkbook.Worksheets(1)\n")
		out.WriteString("    Set cellRef = worksheetRef.Cells((value Mod 10) + 1, 1)\n")
		out.WriteString("    If cellRef Is Nothing Or cellRef.Value = value Then\n")
		for call := 0; call < issue491CallsPerProc; call++ {
			target := (procedure + call + 1) % issue491ProcedureCount
			fmt.Fprintf(&out, "        Call MassiveProcedure%04d(value + %d)\n", target, call+1)
		}
		out.WriteString("        cellRef.Offset(0, 1).Value = value\n")
		out.WriteString("    Else\n")
		out.WriteString("        cellRef.Value = value + 1\n")
		out.WriteString("    End If\n")
		out.WriteString("    On Error Resume Next\n")
		out.WriteString("    Debug.Print cellRef.Address\n")
		out.WriteString("    On Error GoTo 0\n")
		out.WriteString("End Sub\n\n")
	}
	return out.String()
}

func issue491SmallContractSource() string {
	return `Attribute VB_Name = "Issue491Contract"
Option Explicit

Public Sub Reported()
    Dim deck As Collection
    If deck Is Nothing Or deck.Count = 0 Then Exit Sub
End Sub

Public Sub Suppressed()
    Dim deck As Collection
    ' xlflow:disable-next-line VBA212
    If deck Is Nothing Or deck.Count = 0 Then Exit Sub
End Sub
`
}

func TestIssue491SyntheticLargeClassFixture(t *testing.T) {
	first := issue491LargeClassSource()
	second := issue491LargeClassSource()
	if first != second {
		t.Fatal("large class generator is not deterministic")
	}
	if got := strings.Count(first, "Public Sub MassiveProcedure"); got != issue491ProcedureCount {
		t.Fatalf("procedures = %d, want %d", got, issue491ProcedureCount)
	}
	if got := strings.Count(first, "Call MassiveProcedure"); got < 4000 {
		t.Fatalf("calls = %d, want at least 4000", got)
	}
	lines := strings.Count(first, "\n")
	if lines < 20_000 || lines > 25_000 {
		t.Fatalf("lines = %d, want 20,000..25,000", lines)
	}
}

func TestIssue491LightweightProcedureSymbols(t *testing.T) {
	fixture := makeIssue491BenchmarkFixture(t)
	s := newLSPBenchmarkServer(t, fixture)
	doc, err := s.docs.getOrRead(fixture.largeURI)
	if err != nil {
		t.Fatal(err)
	}
	symbols, handled := s.analyzer.LightweightDocumentSymbols(doc, intel.WorkspaceSymbolQuery{Text: "MassiveProcedure0001", Mode: intel.WorkspaceSymbolQueryExact})
	if !handled || len(symbols) == 0 || symbols[0].Name != "MassiveProcedure0001" {
		t.Fatalf("lightweight symbols = %+v, handled=%t", symbols, handled)
	}
	callLine := benchmarkSourceLine(fixture.largeSource, "        Call MassiveProcedure0001(value + 1)")
	var captured intel.WorkspaceSymbolQuery
	original := s.analyzer.WorkspaceSymbolQueryContextFunc
	s.analyzer.WorkspaceSymbolQueryContextFunc = func(ctx context.Context, open []intel.Document, query intel.WorkspaceSymbolQuery) ([]intel.Symbol, error) {
		captured = query
		return original(ctx, open, query)
	}
	_, _ = s.analyzer.SignatureHelp(doc, intel.Position{Line: callLine, Character: 42}, s.docs.openDocuments())
	if captured.Text != "MassiveProcedure0001" {
		t.Fatalf("signature query = %+v, want MassiveProcedure0001", captured)
	}
}

func TestIssue491SmallFixtureDiagnosticContract(t *testing.T) {
	root := t.TempDir()
	path := filepath.Join(root, "Issue491Contract.cls")
	s, cleanup, err := New(Options{RootDir: root, Config: config.Default()})
	if err != nil {
		t.Fatal(err)
	}
	defer cleanup()
	diagnostics := s.analyzer.Diagnostics(intel.Document{
		Path:   path,
		Source: issue491SmallContractSource(),
	})
	if len(diagnostics) != 1 {
		t.Fatalf("diagnostic contract = %+v, want exactly one VBA212 (the other is suppressed)", diagnostics)
	}
	got := diagnostics
	want := intel.Diagnostic{
		Code:     "VBA212",
		Severity: "warning",
		Message:  "deck.Count dereferences deck in the same non-short-circuit boolean expression (OR).",
		Source:   "xlflow",
		Range: intel.Range{
			Start: intel.Position{Line: 5, Character: 25},
			End:   intel.Position{Line: 5, Character: 35},
		},
	}
	if got[0] != want {
		t.Fatalf("VBA212 contract changed:\n got: %+v\nwant: %+v", got[0], want)
	}
	findings, err := analyze.SourceNonShortCircuitObjectGuardFindings(
		root, path, config.Default(), []byte(issue491SmallContractSource()),
	)
	if err != nil {
		t.Fatal(err)
	}
	if len(findings) != 1 || findings[0].Procedure != "Reported" || findings[0].Line != 6 ||
		findings[0].Code != want.Code || findings[0].Severity != want.Severity || findings[0].Message != want.Message {
		t.Fatalf("VBA212 analyzer contract changed: %+v", findings)
	}
}

func TestIssue491DiagnosticsStopsWithinBudgetAfterCancellationCheckpoint(t *testing.T) {
	fixture := makeIssue491BenchmarkFixture(t)
	s := newLSPBenchmarkServer(t, fixture)
	doc, err := s.docs.getOrRead(fixture.largeURI)
	if err != nil {
		t.Fatal(err)
	}
	ctx := &issue491CheckpointContext{cancelAt: 64, reached: make(chan struct{})}
	done := make(chan []intel.Diagnostic, 1)
	go func() { done <- s.analyzer.DiagnosticsContext(ctx, doc) }()
	select {
	case <-ctx.reached:
	case <-time.After(issue491CheckpointWait):
		t.Fatal("diagnostics did not reach a cancellation checkpoint")
	}
	select {
	case diagnostics := <-done:
		if diagnostics != nil {
			t.Fatalf("canceled diagnostics published partial results: %+v", diagnostics)
		}
	case <-time.After(250 * time.Millisecond):
		t.Fatal("obsolete diagnostics did not stop within 250ms of the cancellation checkpoint")
	}
}

func TestIssue494BodyEditReuses1199ProcedureArtifacts(t *testing.T) {
	if testing.Short() {
		t.Skip("large structural artifact-reuse regression")
	}
	fixture := makeIssue491BenchmarkFixture(t)
	s := newLSPBenchmarkServer(t, fixture)
	doc, err := s.docs.getOrRead(fixture.largeURI)
	if err != nil {
		t.Fatal(err)
	}
	_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
	line := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
	column := len("    Debug.Print cellRef.Address")
	changed, applied, err := s.docs.applyChanges(fixture.largeURI, []documentContentChange{{rng: protocolRange(line, column, line, column), text: " "}}, 2)
	if err != nil || !applied {
		t.Fatalf("change = (applied=%v, err=%v)", applied, err)
	}
	_ = s.analyzer.DiagnosticsContext(context.Background(), changed)
	stats := changed.Snapshot.ProcedureArtifactStats()
	if stats.IRBuild != 1 || stats.IRReuse != issue491ProcedureCount-1 || stats.CFGBuild != 1 || stats.CFGReuse != issue491ProcedureCount-1 {
		t.Fatalf("artifact stats = %+v, want one build and %d reuses per artifact", stats, issue491ProcedureCount-1)
	}
}

type issue491CheckpointContext struct {
	checks   atomic.Int64
	cancelAt int64
	reached  chan struct{}
	once     sync.Once
}

func (*issue491CheckpointContext) Deadline() (time.Time, bool) { return time.Time{}, false }
func (*issue491CheckpointContext) Done() <-chan struct{}       { return nil }
func (*issue491CheckpointContext) Value(any) any               { return nil }
func (c *issue491CheckpointContext) Err() error {
	if c.checks.Add(1) < c.cancelAt {
		return nil
	}
	c.once.Do(func() { close(c.reached) })
	return context.Canceled
}

func makeIssue491BenchmarkFixture(tb testing.TB) lspBenchmarkFixture {
	tb.Helper()
	root := tb.TempDir()
	classDir := filepath.Join(root, "src", "classes")
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		tb.Fatal(err)
	}
	source := issue491LargeClassSource()
	path := filepath.Join(classDir, "Issue491LargeClass.cls")
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		tb.Fatal(err)
	}
	return lspBenchmarkFixture{root: root, largePath: path, largeURI: pathToFileURI(path), largeSource: source}
}

func BenchmarkLSPIssue491LargeClass(b *testing.B) {
	fixture := makeIssue491BenchmarkFixture(b)

	b.Run("DidOpenHandler", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
			if err != nil {
				b.Fatal(err)
			}
			ctx := &glsp.Context{Notify: func(string, any) {}}
			b.StartTimer()
			err = s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
				URI: protocol.DocumentUri(fixture.largeURI), Version: int32(i + 1), Text: fixture.largeSource,
			}})
			b.StopTimer()
			cleanup()
			if err != nil {
				b.Fatal(err)
			}
		}
	})

	b.Run("Diagnostics/Cold", func(b *testing.B) {
		b.ReportAllocs()
		var parses uint64
		var fullHashes uint64
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			before := doc.Snapshot.ParseCount()
			beforeHashes := doc.Snapshot.FullHashCount()
			b.StartTimer()
			_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
			b.StopTimer()
			parses += doc.Snapshot.ParseCount() - before
			fullHashes += doc.Snapshot.FullHashCount() - beforeHashes
			cleanup()
		}
		b.ReportMetric(float64(parses)/float64(b.N), "parses/op")
		b.ReportMetric(float64(fullHashes)/float64(b.N), "full-hashes/op")
	})

	b.Run("Diagnostics/Warm", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
		before := doc.Snapshot.ParseCount()
		beforeHashes := doc.Snapshot.FullHashCount()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
		}
		b.ReportMetric(float64(doc.Snapshot.ParseCount()-before)/float64(b.N), "parses/op")
		b.ReportMetric(float64(doc.Snapshot.FullHashCount()-beforeHashes)/float64(b.N), "full-hashes/op")
	})

	b.Run("Diagnostics/EditFast", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		baseline := s.analyzer.DiagnosticsRequestContext(context.Background(), intel.DiagnosticRequest{Document: doc, Mode: intel.DiagnosticModeFull})
		line := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		column := len("    Debug.Print cellRef.Address")
		change := []documentContentChange{{rng: protocolRange(line, column, line, column), text: " "}}
		changed, applied, err := s.docs.applyChanges(fixture.largeURI, change, 2)
		if err != nil || !applied {
			b.Fatalf("change = (applied=%v, err=%v)", applied, err)
		}
		request := intel.DiagnosticRequest{
			Document: changed, Mode: intel.DiagnosticModeFast, PreviousCache: baseline.Cache,
			Changes: intel.ProcedureChangeSet{Ranges: []intel.Range{{Start: intel.Position{Line: line, Character: column}, End: intel.Position{Line: line, Character: column + 1}}}},
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.DiagnosticsRequestContext(context.Background(), request)
		}
	})

	b.Run("Diagnostics/EditFullArtifactReuse", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
		line := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		column := len("    Debug.Print cellRef.Address")
		changed, applied, err := s.docs.applyChanges(fixture.largeURI, []documentContentChange{{rng: protocolRange(line, column, line, column), text: " "}}, 2)
		if err != nil || !applied {
			b.Fatalf("change = (applied=%v, err=%v)", applied, err)
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.DiagnosticsContext(context.Background(), changed)
		}
		stats := changed.Snapshot.ProcedureArtifactStats()
		b.ReportMetric(float64(stats.IRBuild), "ir-builds")
		b.ReportMetric(float64(stats.IRReuse), "ir-reuses")
		b.ReportMetric(float64(stats.CFGBuild), "cfg-builds")
		b.ReportMetric(float64(stats.CFGReuse), "cfg-reuses")
	})

	b.Run("Diagnostics/Canceled", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			ctx, cancel := context.WithCancel(context.Background())
			cancel()
			_ = s.analyzer.DiagnosticsContext(ctx, doc)
		}
	})

	b.Run("WorkspaceQuery/Warm", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		params := &protocol.WorkspaceSymbolParams{Query: "MassiveProcedure1199"}
		if _, err := s.workspaceSymbol(nil, params); err != nil {
			b.Fatal(err)
		}
		beforeBuilds := s.overlayBuilds.Load()
		beforePublications := s.overlayPublications.Load()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := s.workspaceSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
		b.ReportMetric(float64(s.overlayBuilds.Load()-beforeBuilds)/float64(b.N), "overlay-builds/op")
		b.ReportMetric(float64(s.overlayPublications.Load()-beforePublications)/float64(b.N), "overlay-publications/op")
	})

	b.Run("WorkspaceQuery/Cold", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		params := &protocol.WorkspaceSymbolParams{Query: "MassiveProcedure1199"}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			s.analysis = s.newWorkspaceAnalysisIndex()
			if _, err := s.workspaceSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})

	b.Run("SingleCharacterChange", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		line := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		column := len("    Debug.Print cellRef.Address")
		insert := []documentContentChange{{rng: protocolRange(line, column, line, column), text: " "}}
		remove := []documentContentChange{{rng: protocolRange(line, column, line, column+1), text: ""}}
		b.ReportAllocs()
		b.SetBytes(1)
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			changes := insert
			if i%2 == 1 {
				changes = remove
			}
			if _, applied, err := s.docs.applyChanges(fixture.largeURI, changes, int32(i+1)); err != nil || !applied {
				b.Fatalf("change = (applied=%v, err=%v)", applied, err)
			}
		}
	})

	b.Run("ContinuousChanges25", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		line := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		column := len("    Debug.Print cellRef.Address")
		insert := []documentContentChange{{rng: protocolRange(line, column, line, column), text: " "}}
		remove := []documentContentChange{{rng: protocolRange(line, column, line, column+1), text: ""}}
		var version int32
		inserted := false
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			for edit := 0; edit < 25; edit++ {
				changes := insert
				if inserted {
					changes = remove
				}
				version++
				if _, applied, err := s.docs.applyChanges(fixture.largeURI, changes, version); err != nil || !applied {
					b.Fatalf("change %d = (applied=%v, err=%v)", edit, applied, err)
				}
				inserted = !inserted
			}
		}
	})

	benchmarkIssue491Interactive(b, fixture)
	benchmarkIssue491OpenImmediateInteractive(b, fixture)
	benchmarkIssue491Lifecycle(b, fixture)
}

func benchmarkIssue491Interactive(b *testing.B, fixture lspBenchmarkFixture) {
	s := newLSPBenchmarkServer(b, fixture)
	callLine := benchmarkSourceLine(fixture.largeSource, "        Call MassiveProcedure0001(value + 1)")
	memberLine := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")

	b.Run("Interactive/Completion", func(b *testing.B) {
		params := &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(memberLine), Character: 27},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			if _, err := s.completion(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})

	b.Run("Interactive/Hover", func(b *testing.B) {
		params := &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 20},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			if _, err := s.hover(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})

	b.Run("Interactive/SignatureHelp", func(b *testing.B) {
		params := &protocol.SignatureHelpParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 42},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			if _, err := s.signatureHelp(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
}

// interactiveIndexBenchmarkStats keeps the structural counters next to the
// request timing so a benchmark run can distinguish a cold index construction
// from a warm lookup. The values are developer-only observations; they are not
// part of the LSP response.
type interactiveIndexBenchmarkStats struct {
	indexBuilds         uint64
	indexHits           uint64
	catalogBuilds       uint64
	catalogReuses       uint64
	fullSymbolBuilds    uint64
	fullSymbolFallbacks uint64
	parses              uint64
	irBuilds            uint64
	cfgBuilds           uint64
}

func (s interactiveIndexBenchmarkStats) add(other interactiveIndexBenchmarkStats) interactiveIndexBenchmarkStats {
	return interactiveIndexBenchmarkStats{
		indexBuilds:         s.indexBuilds + other.indexBuilds,
		indexHits:           s.indexHits + other.indexHits,
		catalogBuilds:       s.catalogBuilds + other.catalogBuilds,
		catalogReuses:       s.catalogReuses + other.catalogReuses,
		fullSymbolBuilds:    s.fullSymbolBuilds + other.fullSymbolBuilds,
		fullSymbolFallbacks: s.fullSymbolFallbacks + other.fullSymbolFallbacks,
		parses:              s.parses + other.parses,
		irBuilds:            s.irBuilds + other.irBuilds,
		cfgBuilds:           s.cfgBuilds + other.cfgBuilds,
	}
}

func (s interactiveIndexBenchmarkStats) subtract(before interactiveIndexBenchmarkStats) interactiveIndexBenchmarkStats {
	return interactiveIndexBenchmarkStats{
		indexBuilds:         s.indexBuilds - before.indexBuilds,
		indexHits:           s.indexHits - before.indexHits,
		catalogBuilds:       s.catalogBuilds - before.catalogBuilds,
		catalogReuses:       s.catalogReuses - before.catalogReuses,
		fullSymbolBuilds:    s.fullSymbolBuilds - before.fullSymbolBuilds,
		fullSymbolFallbacks: s.fullSymbolFallbacks - before.fullSymbolFallbacks,
		parses:              s.parses - before.parses,
		irBuilds:            s.irBuilds - before.irBuilds,
		cfgBuilds:           s.cfgBuilds - before.cfgBuilds,
	}
}

func interactiveIndexStats(s *Server, doc intel.Document) interactiveIndexBenchmarkStats {
	stats := interactiveIndexBenchmarkStats{}
	if doc.Snapshot != nil && doc.Snapshot.Matches(doc) {
		stats.indexBuilds, stats.indexHits = doc.Snapshot.InteractiveIndexStats()
		stats.catalogBuilds, stats.catalogReuses = doc.Snapshot.ProcedureCatalogStats()
		stats.parses = doc.Snapshot.ParseCount()
		artifacts := doc.Snapshot.ProcedureArtifactStats()
		stats.irBuilds, stats.cfgBuilds = artifacts.IRBuild, artifacts.CFGBuild
	}
	if s != nil && s.performance != nil {
		stats.fullSymbolBuilds = s.performance.counterTotal(performanceCounterFullDocumentSymbolBuilds)
		stats.fullSymbolFallbacks = s.performance.counterTotal(performanceCounterInteractiveFullSymbolFallbacks)
	}
	return stats
}

func reportInteractiveIndexStats(b *testing.B, stats interactiveIndexBenchmarkStats) {
	count := float64(max(1, b.N))
	b.ReportMetric(float64(stats.indexBuilds)/count, "interactive-index-builds/op")
	b.ReportMetric(float64(stats.indexHits)/count, "interactive-index-hits/op")
	b.ReportMetric(float64(stats.catalogBuilds)/count, "procedure-catalog-builds/op")
	b.ReportMetric(float64(stats.catalogReuses)/count, "procedure-catalog-reuses/op")
	b.ReportMetric(float64(stats.fullSymbolBuilds)/count, "full-document-symbol-builds/op")
	b.ReportMetric(float64(stats.fullSymbolFallbacks)/count, "interactive-full-symbol-fallbacks/op")
	b.ReportMetric(float64(stats.parses)/count, "snapshot-parses/op")
	b.ReportMetric(float64(stats.irBuilds)/count, "ir-builds/op")
	b.ReportMetric(float64(stats.cfgBuilds)/count, "cfg-builds/op")
}

type interactiveIndexBenchmarkCase struct {
	name string
	pos  intel.Position
	run  func(intel.Analyzer, intel.Document, []intel.Document, intel.Position) error
}

func runInteractiveIndexBenchmarkCase(s *Server, doc intel.Document, test interactiveIndexBenchmarkCase) error {
	return test.run(s.analyzer, doc, []intel.Document{doc}, test.pos)
}

func benchmarkInteractiveIndexCold(b *testing.B, fixture lspBenchmarkFixture, source string, test interactiveIndexBenchmarkCase) {
	b.Helper()
	s := newLSPBenchmarkServer(b, fixture)
	s.performance = newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	var total interactiveIndexBenchmarkStats
	b.ReportAllocs()
	b.ReportMetric(float64(len(source)), "source-bytes")
	b.ReportMetric(float64(strings.Count(source, "\n")), "source-lines")
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		b.StopTimer()
		snapshot := intel.NewAnalysisSnapshot(intel.Document{
			URI: fixture.largeURI, Path: fixture.largePath, Source: source,
			ModuleKind: benchmarkModuleKind(fixture.largePath), Version: int32(i + 1),
		})
		doc := snapshot.Document()
		before := interactiveIndexStats(s, doc)
		b.StartTimer()
		err := runInteractiveIndexBenchmarkCase(s, doc, test)
		b.StopTimer()
		after := interactiveIndexStats(s, doc)
		total = total.add(after.subtract(before))
		snapshot.Retire()
		if err != nil {
			b.Fatal(err)
		}
	}
	b.StopTimer()
	reportInteractiveIndexStats(b, total)
}

func benchmarkInteractiveIndexWarm(b *testing.B, fixture lspBenchmarkFixture, source string, test interactiveIndexBenchmarkCase) {
	b.Helper()
	s := newLSPBenchmarkServer(b, fixture)
	s.performance = newPerformanceRecorder(true, log.New(io.Discard, "", 0))
	snapshot := intel.NewAnalysisSnapshot(intel.Document{
		URI: fixture.largeURI, Path: fixture.largePath, Source: source,
		ModuleKind: benchmarkModuleKind(fixture.largePath), Version: 1,
	})
	doc := snapshot.Document()
	if err := runInteractiveIndexBenchmarkCase(s, doc, test); err != nil {
		snapshot.Retire()
		b.Fatal(err)
	}
	before := interactiveIndexStats(s, doc)
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if err := runInteractiveIndexBenchmarkCase(s, doc, test); err != nil {
			b.Fatal(err)
		}
	}
	b.StopTimer()
	after := interactiveIndexStats(s, doc)
	snapshot.Retire()
	reportInteractiveIndexStats(b, after.subtract(before))
}

func benchmarkInteractiveIndexCases(b *testing.B, fixture lspBenchmarkFixture, source string, cases []interactiveIndexBenchmarkCase) {
	b.Helper()
	for _, test := range cases {
		b.Run(test.name+"/Cold", func(b *testing.B) {
			benchmarkInteractiveIndexCold(b, fixture, source, test)
		})
		b.Run(test.name+"/Warm", func(b *testing.B) {
			benchmarkInteractiveIndexWarm(b, fixture, source, test)
		})
	}
}

func issue756InteractiveCases(source, lineText, identifier, completionPrefix string) []interactiveIndexBenchmarkCase {
	lineNo := benchmarkSourceLine(source, lineText)
	line := strings.Split(source, "\n")[lineNo]
	start := strings.Index(line, identifier)
	if start < 0 {
		panic("interactive benchmark identifier missing")
	}
	return []interactiveIndexBenchmarkCase{
		{name: "Hover", pos: intel.Position{Line: lineNo, Character: start + 2}, run: func(a intel.Analyzer, doc intel.Document, open []intel.Document, pos intel.Position) error {
			_, err := a.Hover(doc, pos, open)
			return err
		}},
		{name: "Definition", pos: intel.Position{Line: lineNo, Character: start + 2}, run: func(a intel.Analyzer, doc intel.Document, open []intel.Document, pos intel.Position) error {
			_, err := a.Definition(doc, pos, open, pathToFileURI)
			return err
		}},
		{name: "PrefixCompletion", pos: intel.Position{Line: lineNo, Character: start + len(completionPrefix)}, run: func(a intel.Analyzer, doc intel.Document, open []intel.Document, pos intel.Position) error {
			_, err := a.Completions(doc, pos, open)
			return err
		}},
	}
}

func benchmarkModuleKind(path string) string {
	switch strings.ToLower(filepath.Ext(path)) {
	case ".cls":
		return "class"
	case ".frm":
		return "form"
	default:
		return "standard"
	}
}

func firstCallLine(source string) (lineText, identifier string, ok bool) {
	for _, line := range strings.Split(source, "\n") {
		lower := strings.ToLower(line)
		call := strings.Index(lower, "call")
		for call >= 0 {
			if (call == 0 || !isIdentifierByte(line[call-1])) && call+4 < len(line) &&
				(lower[call+4] == ' ' || lower[call+4] == '\t') {
				start := call + 4
				for start < len(line) && (line[start] == ' ' || line[start] == '\t') {
					start++
				}
				end := start
				for end < len(line) && isIdentifierByte(line[end]) {
					end++
				}
				if end > start {
					return line, line[start:end], true
				}
			}
			next := strings.Index(lower[call+4:], "call")
			if next < 0 {
				break
			}
			call += 4 + next
		}
	}
	return "", "", false
}

func isIdentifierByte(value byte) bool {
	return value == '_' || value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value >= '0' && value <= '9'
}

func BenchmarkLSPInteractiveIndexIssue756(b *testing.B) {
	fixture := makeIssue491BenchmarkFixture(b)
	benchmarkInteractiveIndexCases(b, fixture, fixture.largeSource, issue756InteractiveCases(
		fixture.largeSource,
		"        Call MassiveProcedure0001(value + 1)",
		"MassiveProcedure",
		"MassiveProcedure00",
	))
}

func makeIssue756OrdinaryFixture(tb testing.TB) (lspBenchmarkFixture, string) {
	tb.Helper()
	root := tb.TempDir()
	path := filepath.Join(root, "src", "modules", "Interactive.bas")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		tb.Fatal(err)
	}
	source := `Attribute VB_Name = "Interactive"
Option Explicit

Public Function ProjectTarget(ByVal value As Long) As Long
    ProjectTarget = value + 1
End Function

Public Sub Caller()
    Call ProjectTarget(1)
End Sub
`
	if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
		tb.Fatal(err)
	}
	return lspBenchmarkFixture{root: root, largePath: path, largeURI: pathToFileURI(path), largeSource: source}, source
}

func BenchmarkLSPInteractiveIndexOrdinary(b *testing.B) {
	fixture, source := makeIssue756OrdinaryFixture(b)
	benchmarkInteractiveIndexCases(b, fixture, source, issue756InteractiveCases(
		source,
		"    Call ProjectTarget(1)",
		"ProjectTarget",
		"ProjectT",
	))
}

func benchmarkIssue491OpenImmediateInteractive(b *testing.B, fixture lspBenchmarkFixture) {
	callLine := benchmarkSourceLine(fixture.largeSource, "        Call MassiveProcedure0001(value + 1)")
	memberLine := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
	tests := []struct {
		name string
		run  func(*Server) error
	}{
		{name: "Completion", run: func(s *Server) error {
			_, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
				Position:     protocol.Position{Line: protocol.UInteger(memberLine), Character: 27},
			}})
			return err
		}},
		{name: "Hover", run: func(s *Server) error {
			_, err := s.hover(nil, &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
				Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 20},
			}})
			return err
		}},
		{name: "SignatureHelp", run: func(s *Server) error {
			_, err := s.signatureHelp(nil, &protocol.SignatureHelpParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
				Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 42},
			}})
			return err
		}},
	}
	for _, test := range tests {
		b.Run("OpenImmediate/"+test.name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
				if err != nil {
					b.Fatal(err)
				}
				ctx := &glsp.Context{Notify: func(string, any) {}}
				b.StartTimer()
				err = s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
					URI: protocol.DocumentUri(fixture.largeURI), Version: int32(i + 1), Text: fixture.largeSource,
				}})
				if err == nil {
					err = test.run(s)
				}
				b.StopTimer()
				cleanup()
				if err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}

func benchmarkIssue491Lifecycle(b *testing.B, fixture lspBenchmarkFixture) {
	b.Run("Lifecycle/InitialFastDiagnostics", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			state := &diagnosticState{open: true, generation: 1, latest: doc, initialFast: true}
			s.diagWorkers.Add(1)
			b.StartTimer()
			s.runDocumentAnalysis(context.Background(), fixture.largeURI, state, 1, doc, nil, false, intel.DiagnosticModeFast)
			b.StopTimer()
			cleanup()
		}
	})

	b.Run("Lifecycle/DidOpenFirstPublication", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			published := make(chan struct{}, 1)
			ctx := &glsp.Context{Notify: func(method string, _ any) {
				if method != string(protocol.ServerTextDocumentPublishDiagnostics) {
					return
				}
				select {
				case published <- struct{}{}:
				default:
				}
			}}
			b.StartTimer()
			err := s.didOpen(ctx, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
				URI: protocol.DocumentUri(fixture.largeURI), Version: int32(i + 1), Text: fixture.largeSource,
			}})
			if err == nil {
				select {
				case <-published:
				case <-time.After(issue491FirstPublicationWait):
					err = fmt.Errorf("first diagnostics publication timed out")
				}
			}
			b.StopTimer()
			cleanup()
			if err != nil {
				b.Fatal(err)
			}
		}
	})

	b.Run("Lifecycle/FirstFullDiagnostics", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			state := &diagnosticState{open: true, generation: 1, latest: doc}
			s.diagWorkers.Add(1)
			b.StartTimer()
			s.runDocumentAnalysis(context.Background(), fixture.largeURI, state, 1, doc, nil, false, intel.DiagnosticModeFull)
			b.StopTimer()
			cleanup()
		}
	})

	b.Run("Lifecycle/HoverDuringFullDiagnostics", func(b *testing.B) {
		memberLine := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		params := &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(memberLine), Character: protocol.UInteger(len("    Debug.Print "))},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			state := &diagnosticState{open: true, generation: 1, latest: doc}
			started := make(chan struct{}, 1)
			release := make(chan struct{})
			s.performanceHook = func(stage, path string) {
				if path == fixture.largePath && stage == "diagnostics-full-start" {
					select {
					case started <- struct{}{}:
					default:
					}
					<-release
				}
			}
			s.diagWorkers.Add(1)
			done := make(chan struct{})
			go func() {
				s.runDocumentAnalysis(context.Background(), fixture.largeURI, state, 1, doc, nil, false, intel.DiagnosticModeFull)
				close(done)
			}()
			waitLSPStartupEvent(b, started)
			b.StartTimer()
			if _, err := s.hover(nil, params); err != nil {
				b.Fatal(err)
			}
			b.StopTimer()
			close(release)
			select {
			case <-done:
			case <-time.After(30 * time.Second):
				b.Fatal("full diagnostics did not finish")
			}
			cleanup()
		}
	})

	b.Run("Lifecycle/CompletionDuringFullDiagnostics", func(b *testing.B) {
		memberLine := benchmarkSourceLine(fixture.largeSource, "    Debug.Print cellRef.Address")
		params := &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(memberLine), Character: protocol.UInteger(len("    Debug.Print cellRef."))},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			state := &diagnosticState{open: true, generation: 1, latest: doc}
			started := make(chan struct{}, 1)
			release := make(chan struct{})
			s.performanceHook = func(stage, path string) {
				if path == fixture.largePath && stage == "diagnostics-full-start" {
					select {
					case started <- struct{}{}:
					default:
					}
					<-release
				}
			}
			s.diagWorkers.Add(1)
			done := make(chan struct{})
			go func() {
				s.runDocumentAnalysis(context.Background(), fixture.largeURI, state, 1, doc, nil, false, intel.DiagnosticModeFull)
				close(done)
			}()
			waitLSPStartupEvent(b, started)
			b.StartTimer()
			if _, err := s.completion(nil, params); err != nil {
				b.Fatal(err)
			}
			b.StopTimer()
			close(release)
			select {
			case <-done:
			case <-time.After(30 * time.Second):
				b.Fatal("full diagnostics did not finish")
			}
			cleanup()
		}
	})

	b.Run("Lifecycle/DefinitionDuringWorkspaceIndexing", func(b *testing.B) {
		callLine := benchmarkSourceLine(fixture.largeSource, "        Call MassiveProcedure0001(value + 1)")
		params := &protocol.DefinitionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
			TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
			Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 20},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPBenchmarkServerWithCleanup(b, fixture)
			active := make(chan struct{}, 1)
			release := make(chan struct{})
			s.performanceHook = func(stage, path string) {
				if path == fixture.largePath && stage == "declaration-start" {
					select {
					case active <- struct{}{}:
					default:
					}
					<-release
				}
			}
			s.analysis.start()
			waitLSPStartupEvent(b, active)
			b.StartTimer()
			if _, err := s.definition(nil, params); err != nil {
				close(release)
				b.Fatal(err)
			}
			b.StopTimer()
			close(release)
			if err := s.analysis.waitReady(); err != nil {
				b.Fatal(err)
			}
			cleanup()
		}
	})
}

// BenchmarkLSPIssue491OptInROneCOne reads a local large-file specimen only when
// explicitly requested, keeping regular tests and source distributions clean.
// Example:
//
//	XLFLOW_LSP_BENCH_FILE=C:\\...\\ROneCOne.cls
//	XLFLOW_LSP_BENCH_ROOT=C:\\...\\ronecone
func BenchmarkLSPIssue491OptInROneCOne(b *testing.B) {
	path := os.Getenv("XLFLOW_LSP_BENCH_FILE")
	root := os.Getenv("XLFLOW_LSP_BENCH_ROOT")
	if path == "" || root == "" {
		b.Skip("set XLFLOW_LSP_BENCH_FILE and XLFLOW_LSP_BENCH_ROOT to enable the local large-file benchmark")
	}
	body, err := os.ReadFile(path)
	if err != nil {
		b.Fatal(err)
	}
	fixture := lspBenchmarkFixture{
		root: root, largePath: path, largeURI: pathToFileURI(path), largeSource: string(body),
	}
	b.ReportMetric(float64(strings.Count(fixture.largeSource, "\n")), "source_lines")
	b.ReportMetric(float64(len(fixture.largeSource)), "source_bytes")
	if lineText, identifier, ok := firstCallLine(fixture.largeSource); ok {
		prefix := identifier
		if len(prefix) > 2 {
			prefix = prefix[:len(prefix)-2]
		}
		b.Run("Interactive", func(b *testing.B) {
			benchmarkInteractiveIndexCases(b, fixture, fixture.largeSource, issue756InteractiveCases(fixture.largeSource, lineText, identifier, prefix))
		})
	}

	b.Run("Diagnostics/Cold", func(b *testing.B) {
		b.ReportAllocs()
		var parses, fullHashes uint64
		for i := 0; i < b.N; i++ {
			s := newLSPBenchmarkServer(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			beforeParses := doc.Snapshot.ParseCount()
			beforeHashes := doc.Snapshot.FullHashCount()
			_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
			parses += doc.Snapshot.ParseCount() - beforeParses
			fullHashes += doc.Snapshot.FullHashCount() - beforeHashes
		}
		b.ReportMetric(float64(parses)/float64(b.N), "parses/op")
		b.ReportMetric(float64(fullHashes)/float64(b.N), "full-hashes/op")
	})
	b.Run("Diagnostics/Warm", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
		beforeParses := doc.Snapshot.ParseCount()
		beforeHashes := doc.Snapshot.FullHashCount()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.DiagnosticsContext(context.Background(), doc)
		}
		b.ReportMetric(float64(doc.Snapshot.ParseCount()-beforeParses)/float64(b.N), "parses/op")
		b.ReportMetric(float64(doc.Snapshot.FullHashCount()-beforeHashes)/float64(b.N), "full-hashes/op")
	})
	b.Run("WorkspaceQuery/Warm", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		params := &protocol.WorkspaceSymbolParams{Query: "ROneCOne"}
		if _, err := s.workspaceSymbol(nil, params); err != nil {
			b.Fatal(err)
		}
		beforeBuilds := s.overlayBuilds.Load()
		beforePublications := s.overlayPublications.Load()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := s.workspaceSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
		b.ReportMetric(float64(s.overlayBuilds.Load()-beforeBuilds)/float64(b.N), "overlay-builds/op")
		b.ReportMetric(float64(s.overlayPublications.Load()-beforePublications)/float64(b.N), "overlay-publications/op")
	})
}
