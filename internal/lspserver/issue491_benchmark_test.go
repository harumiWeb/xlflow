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

	"github.com/harumiWeb/xlflow/internal/analyze"
	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const (
	issue491ProcedureCount = 1200
	issue491CallsPerProc   = 4
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
		Message:  "deck is guarded against Nothing and dereferenced in the same non-short-circuit boolean expression.",
		Source:   "xlflow",
		Range: intel.Range{
			Start: intel.Position{Line: 5, Character: 0},
			End:   intel.Position{Line: 5, Character: 1},
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
	case <-time.After(5 * time.Second):
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
			s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default()})
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
			s := newLSPBenchmarkServer(b, fixture)
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
			Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 46},
		}}
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			if _, err := s.signatureHelp(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
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
				Position:     protocol.Position{Line: protocol.UInteger(callLine), Character: 46},
			}})
			return err
		}},
	}
	for _, test := range tests {
		b.Run("OpenImmediate/"+test.name, func(b *testing.B) {
			b.ReportAllocs()
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default()})
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
