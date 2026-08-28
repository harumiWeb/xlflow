package lspserver

import (
	"context"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"sync"
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
	s, cleanup := newLSPBenchmarkServerWithCleanup(tb, fixture)
	tb.Cleanup(cleanup)
	return s
}

func newLSPBenchmarkServerWithCleanup(tb testing.TB, fixture lspBenchmarkFixture) (*Server, func()) {
	tb.Helper()
	s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
	if err != nil {
		tb.Fatal(err)
	}
	if _, err := s.docs.open(fixture.largeURI, fixture.largeSource); err != nil {
		cleanup()
		tb.Fatal(err)
	}
	return s, cleanup
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
			s.analysis = s.newWorkspaceAnalysisIndex()
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
		var parses uint64
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s := newLSPBenchmarkServer(b, fixture)
			doc, err := s.docs.getOrRead(fixture.largeURI)
			if err != nil {
				b.Fatal(err)
			}
			before := doc.Snapshot.ParseCount()
			b.StartTimer()
			_ = s.analyzer.Diagnostics(doc)
			parses += doc.Snapshot.ParseCount() - before
		}
		b.ReportMetric(float64(parses)/float64(b.N), "parses/op")
	})
	b.Run("Repeated", func(b *testing.B) {
		s := newLSPBenchmarkServer(b, fixture)
		doc, err := s.docs.getOrRead(fixture.largeURI)
		if err != nil {
			b.Fatal(err)
		}
		_ = s.analyzer.Diagnostics(doc)
		before := doc.Snapshot.ParseCount()
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			_ = s.analyzer.Diagnostics(doc)
		}
		b.ReportMetric(float64(doc.Snapshot.ParseCount()-before)/float64(b.N), "parses/op")
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

// BenchmarkLSPIncrementalSingleCharacterEdit keeps the immutable snapshot
// publication cost visible while reporting the one-byte incremental payload
// against the full-document payload used by BenchmarkLSPEdit.
func BenchmarkLSPIncrementalSingleCharacterEdit(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	lineNo := benchmarkSourceLine(fixture.largeSource, "    localSheet.Ra")
	line := strings.Split(fixture.largeSource, "\n")[lineNo]
	column := strings.Index(line, "Ra") + len("Ra")
	if column < len("Ra") {
		b.Fatal("benchmark edit target missing")
	}
	insert := []documentContentChange{{rng: protocolRange(lineNo, column, lineNo, column), text: "n"}}
	remove := []documentContentChange{{rng: protocolRange(lineNo, column, lineNo, column+1), text: ""}}
	b.ReportAllocs()
	b.SetBytes(1)
	for i := 0; i < b.N; i++ {
		changes := insert
		if i%2 == 1 {
			changes = remove
		}
		if _, applied, err := s.docs.applyChanges(fixture.largeURI, changes, int32(i+1)); err != nil || !applied {
			b.Fatalf("incremental change = (applied=%v, err=%v)", applied, err)
		}
	}
	b.ReportMetric(float64(len(fixture.largeSource)), "full-change-bytes/op")
	b.ReportMetric(1, "parses/op")
}

// BenchmarkLSPFullParseSingleCharacterEdit is the comparable complete-parse
// baseline for BenchmarkLSPIncrementalSingleCharacterEdit.
func BenchmarkLSPFullParseSingleCharacterEdit(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	changed := strings.Replace(fixture.largeSource, "    localSheet.Ra\n", "    localSheet.Ran\n", 1)
	b.ReportAllocs()
	b.SetBytes(int64(len(fixture.largeSource)))
	var parses uint64
	for i := 0; i < b.N; i++ {
		source := fixture.largeSource
		if i%2 == 0 {
			source = changed
		}
		doc, err := s.docs.change(fixture.largeURI, source, int32(i+1))
		if err != nil {
			b.Fatal(err)
		}
		before := doc.Snapshot.ParseCount()
		if _, err := doc.Snapshot.ParsedDocument(); err != nil {
			b.Fatal(err)
		}
		parses += doc.Snapshot.ParseCount() - before
	}
	b.ReportMetric(float64(parses)/float64(b.N), "parses/op")
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
		var runs atomic.Int64
		var discarded atomic.Int64
		var published atomic.Int64
		s.diagnostics = func(ctx context.Context, doc intel.Document) []intel.Diagnostic {
			runs.Add(1)
			current := active.Add(1)
			for {
				observed := maximum.Load()
				if current <= observed || maximum.CompareAndSwap(observed, current) {
					break
				}
			}
			started <- doc.Version
			<-release
			if ctx.Err() != nil {
				discarded.Add(1)
			}
			active.Add(-1)
			completed <- doc.Version
			return nil
		}
		ctx := &glsp.Context{Notify: func(string, any) { published.Add(1) }}
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
					select {
					case <-settle.C:
					default:
					}
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
		b.ReportMetric(float64(runs.Load()), "diagnostic_runs")
		b.ReportMetric(float64(discarded.Load()), "discarded_runs")
		b.ReportMetric(float64(published.Load()), "published_runs")
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
	b.Run("Cold", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			s.analysis = s.newWorkspaceAnalysisIndex()
			s.semanticTokens = newSemanticTokenCache()
			if _, err := s.docs.change(fixture.largeURI, fixture.largeSource, int32(i+1)); err != nil {
				b.Fatal(err)
			}
			if _, err := s.semanticTokensFull(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
	b.Run("Warm", func(b *testing.B) {
		if _, err := s.semanticTokensFull(nil, params); err != nil {
			b.Fatal(err)
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := s.semanticTokensFull(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
	b.Run("AfterEdit", func(b *testing.B) {
		changed := strings.Replace(fixture.largeSource, "    localSheet.Ra\n", "    localSheet.Ran\n", 1)
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			source := fixture.largeSource
			if i%2 == 0 {
				source = changed
			}
			if _, err := s.docs.change(fixture.largeURI, source, int32(i+1)); err != nil {
				b.Fatal(err)
			}
			s.semanticTokens.invalidateAll()
			if _, err := s.semanticTokensFull(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
}

func BenchmarkLSPDocumentSymbols(b *testing.B) {
	fixture := makeLSPBenchmarkFixture(b)
	s := newLSPBenchmarkServer(b, fixture)
	params := &protocol.DocumentSymbolParams{TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.largeURI)}}
	b.Run("Cold", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			if _, err := s.docs.change(fixture.largeURI, fixture.largeSource, int32(i+1)); err != nil {
				b.Fatal(err)
			}
			if _, err := s.documentSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
	b.Run("Warm", func(b *testing.B) {
		if _, err := s.documentSymbol(nil, params); err != nil {
			b.Fatal(err)
		}
		b.ReportAllocs()
		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			if _, err := s.documentSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
	b.Run("AfterEdit", func(b *testing.B) {
		changed := strings.Replace(fixture.largeSource, "    localSheet.Ra\n", "    localSheet.Ran\n", 1)
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			source := fixture.largeSource
			if i%2 == 0 {
				source = changed
			}
			if _, err := s.docs.change(fixture.largeURI, source, int32(i+1)); err != nil {
				b.Fatal(err)
			}
			if _, err := s.documentSymbol(nil, params); err != nil {
				b.Fatal(err)
			}
		}
	})
}

func benchmarkSourceLine(source, exact string) int {
	for line, text := range strings.Split(source, "\n") {
		if text == exact {
			return line
		}
	}
	panic("benchmark source line not found: " + exact)
}

type lspStartupBenchmarkFixture struct {
	root       string
	moduleA    string
	moduleB    string
	moduleAURI string
	moduleBURI string
	sourceA    string
}

func makeLSPStartupBenchmarkFixture(tb testing.TB) lspStartupBenchmarkFixture {
	tb.Helper()
	root := tb.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		tb.Fatal(err)
	}
	moduleB := filepath.Join(moduleDir, "00_ModuleB.bas")
	moduleA := filepath.Join(moduleDir, "01_ModuleA.bas")
	sourceB := "Attribute VB_Name = \"ModuleB\"\nOption Explicit\n\nPublic Function CrossFileTarget(ByVal value As Long) As Long\n    CrossFileTarget = value + 1\nEnd Function\n"
	sourceA := "Attribute VB_Name = \"ModuleA\"\nOption Explicit\n\nPublic Sub CallB()\n    Call CrossFileTarget(1)\nEnd Sub\n\nPublic Sub CompletionTarget()\n    Worksheets(\"Input\").Ra\nEnd Sub\n"
	if err := os.WriteFile(moduleB, []byte(sourceB), 0o644); err != nil {
		tb.Fatal(err)
	}
	if err := os.WriteFile(moduleA, []byte(sourceA), 0o644); err != nil {
		tb.Fatal(err)
	}
	for i := 0; i < 32; i++ {
		path := filepath.Join(moduleDir, fmt.Sprintf("%02d_Filler.bas", i+2))
		source := fmt.Sprintf("Attribute VB_Name = \"Filler%02d\"\nOption Explicit\n\nPublic Sub FillerProcedure%02d()\nEnd Sub\n", i, i)
		if err := os.WriteFile(path, []byte(source), 0o644); err != nil {
			tb.Fatal(err)
		}
	}
	return lspStartupBenchmarkFixture{
		root: root, moduleA: moduleA, moduleB: moduleB,
		moduleAURI: pathToFileURI(moduleA), moduleBURI: pathToFileURI(moduleB), sourceA: sourceA,
	}
}

func TestLSPStartupBenchmarkFixture(t *testing.T) {
	first := makeLSPStartupBenchmarkFixture(t)
	second := makeLSPStartupBenchmarkFixture(t)
	if first.sourceA != second.sourceA {
		t.Fatal("startup fixture is not deterministic")
	}
	files, err := filepath.Glob(filepath.Join(first.root, "src", "modules", "*.bas"))
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 34 {
		t.Fatalf("startup fixture files = %d, want 34", len(files))
	}
}

func startupInteractiveQueries(b *testing.B, s *Server, fixture lspStartupBenchmarkFixture) (int, int, int) {
	b.Helper()
	line := benchmarkSourceLine(fixture.sourceA, "    Call CrossFileTarget(1)")
	character := len("    Call ")
	hover, err := s.hover(nil, &protocol.HoverParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(character)},
	}})
	if err != nil {
		b.Fatal(err)
	}
	definition, err := s.definition(nil, &protocol.DefinitionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(line), Character: protocol.UInteger(character)},
	}})
	if err != nil {
		b.Fatal(err)
	}
	definitions := 0
	if locations, ok := definition.([]protocol.Location); ok {
		for _, location := range locations {
			if location.URI == protocol.DocumentUri(fixture.moduleBURI) {
				definitions++
			}
		}
	}
	hovers := 0
	if hover != nil {
		if contents, ok := hover.Contents.(protocol.MarkupContent); ok && strings.Contains(contents.Value, "CrossFileTarget") {
			hovers = 1
		}
	}
	completionLine := benchmarkSourceLine(fixture.sourceA, "    Worksheets(\"Input\").Ra")
	completion, err := s.completion(nil, &protocol.CompletionParams{TextDocumentPositionParams: protocol.TextDocumentPositionParams{
		TextDocument: protocol.TextDocumentIdentifier{URI: protocol.DocumentUri(fixture.moduleAURI)},
		Position:     protocol.Position{Line: protocol.UInteger(completionLine), Character: 27},
	}})
	if err != nil {
		b.Fatal(err)
	}
	completions := 0
	if list, ok := completion.(protocol.CompletionList); ok && len(list.Items) > 0 {
		completions = 1
	}
	return hovers, definitions, completions
}

func waitLSPStartupEvent(b *testing.B, events <-chan struct{}) {
	b.Helper()
	select {
	case <-events:
	case <-time.After(30 * time.Second):
		b.Fatal("timed out waiting for startup benchmark checkpoint")
	}
}

func startLSPStartupServer(b *testing.B, fixture lspStartupBenchmarkFixture) (*Server, func()) {
	b.Helper()
	s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
	if err != nil {
		b.Fatal(err)
	}
	if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
		cleanup()
		b.Fatal(err)
	}
	if err := s.initialized(nil, nil); err != nil {
		cleanup()
		b.Fatal(err)
	}
	if err := s.didOpen(nil, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
		URI: protocol.DocumentUri(fixture.moduleAURI), Version: 1, Text: fixture.sourceA,
	}}); err != nil {
		cleanup()
		b.Fatal(err)
	}
	return s, cleanup
}

func newLSPStartupServer(b *testing.B, fixture lspStartupBenchmarkFixture) (*Server, func()) {
	b.Helper()
	s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
	if err != nil {
		b.Fatal(err)
	}
	return s, cleanup
}

// BenchmarkLSPStartup reproduces the initialize/initialized/background scan/
// didOpen/interactive overlap. Checkpoints are synchronized by internal
// parser hooks rather than wall-clock sleeps, so the benchmark remains useful
// on contended machines and makes pre-readiness behavior explicit.
func BenchmarkLSPStartup(b *testing.B) {
	fixture := makeLSPStartupBenchmarkFixture(b)

	b.Run("Initialization", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup, err := New(Options{RootDir: fixture.root, Config: config.Default(), Stderr: io.Discard})
			if err != nil {
				b.Fatal(err)
			}
			b.StartTimer()
			if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
				b.Fatal(err)
			}
			if err := s.initialized(nil, nil); err != nil {
				b.Fatal(err)
			}
			if err := s.analysis.waitReady(); err != nil {
				b.Fatal(err)
			}
			b.StopTimer()
			cleanup()
		}
	})

	b.Run("ImmediateInteractive", func(b *testing.B) {
		b.ReportAllocs()
		var hovers, definitions, completions int
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := startLSPStartupServer(b, fixture)
			b.StartTimer()
			hover, definition, completion := startupInteractiveQueries(b, s, fixture)
			hovers += hover
			definitions += definition
			completions += completion
			b.StopTimer()
			if err := s.analysis.waitReady(); err != nil {
				b.Fatal(err)
			}
			cleanup()
		}
		b.ReportMetric(float64(hovers)/float64(b.N), "hover_results/op")
		b.ReportMetric(float64(definitions)/float64(b.N), "definition_results/op")
		b.ReportMetric(float64(completions)/float64(b.N), "completion_results/op")
	})

	b.Run("WhileIndexing", func(b *testing.B) {
		b.ReportAllocs()
		var hovers, definitions, completions int
		for i := 0; i < b.N; i++ {
			func() {
				b.StopTimer()
				s, cleanup := newLSPStartupServer(b, fixture)
				defer cleanup()
				active := make(chan struct{}, 1)
				releaseCh := make(chan struct{})
				var releaseOnce sync.Once
				release := func() { releaseOnce.Do(func() { close(releaseCh) }) }
				defer release()
				s.performanceHook = func(stage, path string) {
					if path == fixture.moduleB && stage == "declaration-start" {
						select {
						case active <- struct{}{}:
						default:
						}
						<-releaseCh
					}
				}
				if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
					b.Fatal(err)
				}
				if err := s.initialized(nil, nil); err != nil {
					b.Fatal(err)
				}
				if err := s.didOpen(nil, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
					URI: protocol.DocumentUri(fixture.moduleAURI), Version: 1, Text: fixture.sourceA,
				}}); err != nil {
					b.Fatal(err)
				}
				waitLSPStartupEvent(b, active)
				b.StartTimer()
				hover, definition, completion := startupInteractiveQueries(b, s, fixture)
				hovers += hover
				definitions += definition
				completions += completion
				b.StopTimer()
				release()
				if err := s.analysis.waitReady(); err != nil {
					b.Fatal(err)
				}
			}()
		}
		b.ReportMetric(float64(hovers)/float64(b.N), "hover_results/op")
		b.ReportMetric(float64(definitions)/float64(b.N), "definition_results/op")
		b.ReportMetric(float64(completions)/float64(b.N), "completion_results/op")
	})

	b.Run("AfterDeclarationBeforeSemantic", func(b *testing.B) {
		b.ReportAllocs()
		var hovers, definitions, completions int
		for i := 0; i < b.N; i++ {
			func() {
				b.StopTimer()
				s, cleanup := newLSPStartupServer(b, fixture)
				defer cleanup()
				declared := make(chan struct{}, 1)
				releaseCh := make(chan struct{})
				var releaseOnce sync.Once
				release := func() { releaseOnce.Do(func() { close(releaseCh) }) }
				defer release()
				s.performanceHook = func(stage, path string) {
					if path != fixture.moduleB {
						return
					}
					switch stage {
					case "declaration-ready":
						select {
						case declared <- struct{}{}:
						default:
						}
					case "semantic-start":
						<-releaseCh
					}
				}
				if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
					b.Fatal(err)
				}
				if err := s.initialized(nil, nil); err != nil {
					b.Fatal(err)
				}
				if err := s.didOpen(nil, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
					URI: protocol.DocumentUri(fixture.moduleAURI), Version: 1, Text: fixture.sourceA,
				}}); err != nil {
					b.Fatal(err)
				}
				waitLSPStartupEvent(b, declared)
				// The parser hook marks the declaration parse boundary; wait for the
				// corresponding posting publication before measuring the query itself.
				waitForWorkspaceSymbol(b, s.analysis, "CrossFileTarget")
				b.StartTimer()
				hover, definition, completion := startupInteractiveQueries(b, s, fixture)
				if hover != 1 || definition != 1 {
					b.Fatalf("declaration-ready interactive results = hover %d, definition %d; want 1/1", hover, definition)
				}
				hovers += hover
				definitions += definition
				completions += completion
				b.StopTimer()
				release()
				if err := s.analysis.waitReady(); err != nil {
					b.Fatal(err)
				}
			}()
		}
		b.ReportMetric(float64(hovers)/float64(b.N), "hover_results/op")
		b.ReportMetric(float64(definitions)/float64(b.N), "definition_results/op")
		b.ReportMetric(float64(completions)/float64(b.N), "completion_results/op")
	})

	b.Run("AfterSemanticReady", func(b *testing.B) {
		b.ReportAllocs()
		var hovers, definitions, completions int
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			s, cleanup := newLSPStartupServer(b, fixture)
			semantic := make(chan struct{}, 1)
			s.performanceHook = func(stage, path string) {
				if path == fixture.moduleB && stage == "semantic-ready" {
					select {
					case semantic <- struct{}{}:
					default:
					}
				}
			}
			if _, err := s.initialize(nil, &protocol.InitializeParams{}); err != nil {
				b.Fatal(err)
			}
			if err := s.initialized(nil, nil); err != nil {
				b.Fatal(err)
			}
			if err := s.didOpen(nil, &protocol.DidOpenTextDocumentParams{TextDocument: protocol.TextDocumentItem{
				URI: protocol.DocumentUri(fixture.moduleAURI), Version: 1, Text: fixture.sourceA,
			}}); err != nil {
				b.Fatal(err)
			}
			waitLSPStartupEvent(b, semantic)
			if err := s.analysis.waitReady(); err != nil {
				b.Fatal(err)
			}
			b.StartTimer()
			hover, definition, completion := startupInteractiveQueries(b, s, fixture)
			hovers += hover
			definitions += definition
			completions += completion
			b.StopTimer()
			cleanup()
		}
		b.ReportMetric(float64(hovers)/float64(b.N), "hover_results/op")
		b.ReportMetric(float64(definitions)/float64(b.N), "definition_results/op")
		b.ReportMetric(float64(completions)/float64(b.N), "completion_results/op")
	})
}
