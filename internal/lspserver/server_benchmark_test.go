package lspserver

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync/atomic"
	"testing"
	"time"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vba/intel"
	"github.com/tliron/glsp"
	protocol "github.com/tliron/glsp/protocol_3_16"
)

const (
	benchmarkProjectModules    = 100
	benchmarkProceduresPerFile = 20
	benchmarkLargeProcedures   = 24
)

type lspBenchmarkFixture struct {
	root        string
	largePath   string
	largeURI    string
	largeSource string
}

func makeLSPBenchmarkFixture(tb testing.TB) lspBenchmarkFixture {
	tb.Helper()
	root := tb.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	classDir := filepath.Join(root, "src", "classes")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		tb.Fatal(err)
	}
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		tb.Fatal(err)
	}

	for module := 0; module < benchmarkProjectModules; module++ {
		ext := ".bas"
		dir := moduleDir
		kind := "Module"
		if module%4 == 0 {
			ext = ".cls"
			dir = classDir
			kind = "Class"
		}
		name := fmt.Sprintf("Project%s%03d", kind, module)
		source := projectModuleSource(name, module)
		if err := os.WriteFile(filepath.Join(dir, name+ext), []byte(source), 0o644); err != nil {
			tb.Fatal(err)
		}
	}

	largePath := filepath.Join(moduleDir, "LargeModule.bas")
	largeSource := largeModuleSource()
	if err := os.WriteFile(largePath, []byte(largeSource), 0o644); err != nil {
		tb.Fatal(err)
	}
	return lspBenchmarkFixture{
		root:        root,
		largePath:   largePath,
		largeURI:    pathToFileURI(largePath),
		largeSource: largeSource,
	}
}

func projectModuleSource(name string, module int) string {
	var out strings.Builder
	fmt.Fprintf(&out, "Attribute VB_Name = %q\nOption Explicit\n\n", name)
	for procedure := 0; procedure < benchmarkProceduresPerFile; procedure++ {
		fmt.Fprintf(&out, "' Project procedure %03d/%02d.\n", module, procedure)
		fmt.Fprintf(&out, "Public Function ProjectFunction%03d_%02d(ByVal value As Long) As Long\n", module, procedure)
		fmt.Fprintf(&out, "    ProjectFunction%03d_%02d = value + %d\n", module, procedure, module+procedure)
		out.WriteString("End Function\n\n")
	}
	return out.String()
}

func largeModuleSource() string {
	var out strings.Builder
	out.WriteString("Attribute VB_Name = \"LargeModule\"\nOption Explicit\n\n")
	out.WriteString("Private moduleCounter As Long\n\n")
	for procedure := 0; procedure < benchmarkLargeProcedures; procedure++ {
		fmt.Fprintf(&out, "' Computes a deterministic worksheet value for procedure %03d.\n", procedure)
		fmt.Fprintf(&out, "' @param inputValue Value to transform in procedure %03d.\n", procedure)
		fmt.Fprintf(&out, "' @return The transformed value for procedure %03d.\n", procedure)
		fmt.Fprintf(&out, "Public Function LargeProcedure%03d(ByVal inputValue As Long, Optional ByVal factor As Long = 2) As Long\n", procedure)
		out.WriteString("    Dim sheet As Worksheet\n")
		out.WriteString("    Dim cell As Range\n")
		out.WriteString("    Set sheet = ThisWorkbook.Worksheets(1)\n")
		out.WriteString("    With sheet\n")
		out.WriteString("        With .Cells(1, 1)\n")
		out.WriteString("            Set cell = .Offset(RowOffset:=inputValue Mod 3, ColumnOffset:=factor Mod 3)\n")
		out.WriteString("            cell.Font.Bold = (inputValue > 0)\n")
		out.WriteString("        End With\n")
		out.WriteString("    End With\n")
		fmt.Fprintf(&out, "    moduleCounter = inputValue + %d\n", procedure)
		fmt.Fprintf(&out, "    LargeProcedure%03d = moduleCounter + factor\n", procedure)
		out.WriteString("End Function\n\n")
	}
	out.WriteString("Public Sub CompletionTarget()\n")
	out.WriteString("    Dim localSheet As Worksheet\n")
	out.WriteString("    Set localSheet = ThisWorkbook.Worksheets(1)\n")
	out.WriteString("    localSheet.Ra\n")
	out.WriteString("    LargeProcedure000 1, factor:=3\n")
	out.WriteString("    ProjectFunction005_00 1\n")
	out.WriteString("End Sub\n")
	return out.String()
}

func newLSPBenchmarkServer(tb testing.TB, fixture lspBenchmarkFixture) *Server {
	tb.Helper()
	s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default()})
	if err != nil {
		tb.Fatal(err)
	}
	tb.Cleanup(cleanup)
	if _, err := s.docs.open(fixture.largeURI, fixture.largeSource); err != nil {
		tb.Fatal(err)
	}
	return s
}

func TestLSPBenchmarkFixture(t *testing.T) {
	first := makeLSPBenchmarkFixture(t)
	second := makeLSPBenchmarkFixture(t)
	if first.largeSource != second.largeSource {
		t.Fatal("large module generator is not deterministic")
	}
	if got := strings.Count(first.largeSource, "Public Function LargeProcedure"); got != benchmarkLargeProcedures {
		t.Fatalf("large procedures = %d, want %d", got, benchmarkLargeProcedures)
	}
	for _, fixture := range []lspBenchmarkFixture{first, second} {
		files := 0
		declarations := 0
		err := filepath.WalkDir(filepath.Join(fixture.root, "src"), func(path string, entry os.DirEntry, err error) error {
			if err != nil {
				return err
			}
			if entry.IsDir() || path == fixture.largePath {
				return nil
			}
			body, err := os.ReadFile(path)
			if err != nil {
				return err
			}
			files++
			declarations += strings.Count(string(body), "Public Function ProjectFunction")
			return nil
		})
		if err != nil {
			t.Fatal(err)
		}
		if files != benchmarkProjectModules {
			t.Fatalf("project files = %d, want %d", files, benchmarkProjectModules)
		}
		if declarations != benchmarkProjectModules*benchmarkProceduresPerFile {
			t.Fatalf("project declarations = %d, want %d", declarations, benchmarkProjectModules*benchmarkProceduresPerFile)
		}
	}
}

func BenchmarkLSPWorkspaceSymbols(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	params := &protocol.WorkspaceSymbolParams{Query: "ProjectFunction"}

	b.Run("Cold", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			s.symbols = newWorkspaceSymbolCache()
			if _, err := s.workspaceSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
	b.Run("Warm", func(b *testing.B) {
		if _, err := s.workspaceSymbol(nil, params); err != nil {
			b.Fatal(err)
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := s.workspaceSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
}

func BenchmarkLSPDiagnostics(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)

	b.Run("First", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s := newLSPBenchmarkServer(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			b.StartTimer()
			_ = s.analyzer.Diagnostics(doc)
		}
	})
	b.Run("Repeated", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		_ = s.analyzer.Diagnostics(doc)
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.Diagnostics(doc)
		}
	})
}

func BenchmarkLSPEdit(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	changed := strings.Replace(fixture.largeSource, "    localSheet.Ra\n", "    localSheet.Ran\n", 1)

	b.Run("Single", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			source := fixture.largeSource
			if i%2 == 0 {
				source = changed
			}
			doc, err := s.docs.change(fixture.largeURI, source)
			if err != nil {
				b.Fatal(err)
			}
			_ = s.analyzer.Diagnostics(doc)
		}
	})
	b.Run("Continuous25", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			var doc intel.Document
			for edit := 0; edit < 25; edit++ {
				source := fixture.largeSource
				if edit%2 == 0 {
					source = changed
				}
				var err error
				doc, err = s.docs.change(fixture.largeURI, source)
				if err != nil {
					b.Fatal(err)
				}
			}
			_ = s.analyzer.Diagnostics(doc)
		}
	})
}

func BenchmarkLSPContinuousEditScheduling(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	const edits = 25
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		b.StopTimer()
		s := newLSPBenchmarkServer(b, fixture)
		s.diagnosticsDebounce = 0
		started := make(chan int32, edits)
		completed := make(chan int32, edits)
		release := make(chan struct{})
		var active atomic.Int64
		var maximum atomic.Int64
		s.diagnostics = func(doc intel.Document) []intel.Diagnostic {
			current := active.Add(1)
			for {
				observed := maximum.Load()
				if current <= observed || maximum.CompareAndSwap(observed, current) {
					break
				}
			}
			started <- doc.Version
			<-release
			active.Add(-1)
			completed <- doc.Version
			return nil
		}
		ctx := &glsp.Context{Notify: func(string, any) {}}
		b.StartTimer()
		for edit := 0; edit < edits; edit++ {
			source := fixture.largeSource
			if edit%2 == 0 {
				source = strings.Replace(source, "    localSheet.Ra\n", "    localSheet.Ran\n", 1)
			}
			err := s.didChange(ctx, &protocol.DidChangeTextDocumentParams{
				TextDocument: protocol.VersionedTextDocumentIdentifier{
					TextDocumentIdentifier: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
					Version:                int32(edit + 1),
				},
				ContentChanges: []any{protocol.TextDocumentContentChangeEventWhole{Text: source}},
			})
			if err != nil {
				b.Fatal(err)
			}
			if edit == 0 {
				<-started
			}
		}
		b.StopTimer()
		settle := time.NewTimer(50 * time.Millisecond)
		deadline := time.NewTimer(2 * time.Second)
	settling:
		for maximum.Load() < edits {
			select {
			case <-started:
				if !settle.Stop() {
					<-settle.C
				}
				settle.Reset(50 * time.Millisecond)
			case <-settle.C:
				break settling
			case <-deadline.C:
				break settling
			}
		}
		if !settle.Stop() {
			select {
			case <-settle.C:
			default:
			}
		}
		if !deadline.Stop() {
			select {
			case <-deadline.C:
			default:
			}
		}
		close(release)
		completionDeadline := time.NewTimer(5 * time.Second)
		latestCompleted := false
		for {
			select {
			case version := <-completed:
				if version == edits {
					latestCompleted = true
				}
				if latestCompleted && active.Load() == 0 {
					completionDeadline.Stop()
					goto complete
				}
			case <-completionDeadline.C:
				b.Fatal("latest diagnostics generation did not complete")
			}
		}
	complete:
		b.ReportMetric(float64(maximum.Load()), "max_concurrent")
	}
}

func BenchmarkLSPCompletion(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	line := benchmarkSourceLine(fixture.largeSource, "    localSheet.Ra")
	params := &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
		Position:     protocol.Position{Line: protocol.UInteger(line), Character: 17},
	}}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := s.completion(nil, params); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkLSPHover(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	cases := []struct {
		name      string
		lineText  string
		character protocol.UInteger
	}{
		{name: "Local", lineText: "    localSheet.Ra", character: 10},
		{name: "Project", lineText: "    ProjectFunction005_00 1", character: 12},
	}
	for _, tc := range cases {
		b.Run(tc.name, func(b *testing.B) {
			params := &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
				TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)},
				Position: protocol.Position{
					Line:      protocol.UInteger(benchmarkSourceLine(fixture.largeSource, tc.lineText)),
					Character: tc.character,
				},
			}}
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if _, err := s.hover(nil, params); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}

func BenchmarkLSPSemanticTokens(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	params := &protocol.SemanticTokensParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)}}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := s.semanticTokensFull(nil, params); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkLSPDocumentSymbols(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	params := &protocol.DocumentSymbolParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)}}
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := s.documentSymbol(nil, params); err != nil {
			b.Fatal(err)
		}
	}
}

func benchmarkSourceLine(source, exact string) int {
	for line, text := range strings.Split(source, "\n") {
		if text == exact {
			return line
		}
	}
	panic("benchmark source line not found: " + exact)
}
