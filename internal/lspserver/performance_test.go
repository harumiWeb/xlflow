package lspserver

import (
	"bytes"
	"errors"
	"strings"
	"testing"

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

func TestWorkspaceSymbolCachePerformanceReportsMissThenHit(t *testing.T) {
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
	if _, err := s.cachedWorkspaceSymbols(nil, ""); err != nil {
		t.Fatal(err)
	}
	logOutput := output.String()
	if !strings.Contains(logOutput, `operation="workspaceSymbols/cache/base"`) ||
		!strings.Contains(logOutput, `cache="miss"`) ||
		!strings.Contains(logOutput, `cache="hit"`) {
		t.Fatalf("cache performance log missing miss/hit events:\n%s", logOutput)
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
