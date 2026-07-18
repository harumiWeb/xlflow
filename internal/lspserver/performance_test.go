package lspserver

import (
	"bytes"
	"errors"
	"path/filepath"
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
	logOutput := output.String()
	if !strings.Contains(logOutput, `operation="workspaceSymbols/index/initial"`) ||
		!strings.Contains(logOutput, `file_count=0`) ||
		!strings.Contains(logOutput, `elapsed_ms=`) {
		t.Fatalf("workspace index performance log missing initial build fields:\n%s", logOutput)
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
	s.semanticTokenGenerator = func(intel.Document, []intel.Document) ([]intel.SemanticToken, error) {
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
