package intel

import (
	"errors"
	"reflect"
	"sync"
	"sync/atomic"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/calls"
	vbacfg "github.com/harumiWeb/xlflow/internal/vba/cfg"
	"github.com/harumiWeb/xlflow/internal/vba/doccomments"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
	"github.com/harumiWeb/xlflow/internal/vba/symbols"
	tree_sitter "github.com/tree-sitter/go-tree-sitter"
)

func TestAnalysisSnapshotIdentityLinesAndProcedures(t *testing.T) {
	doc := Document{
		URI: "file:///C:/work/Main.bas", Path: `C:\work\Main.bas`, Version: 7,
		ModuleKind: "standard",
		Source:     "Option Explicit\r\nPublic Sub First()\r  Dim value As Long\nEnd Sub\nPrivate Function Second() As Long\nEnd Function\n",
	}
	snapshot := NewAnalysisSnapshot(doc)
	view := snapshot.Document()
	if view.Snapshot != snapshot || view.Version != 7 || view.Source != doc.Source {
		t.Fatalf("snapshot document = %+v", view)
	}
	if got := snapshot.SourceHash(); len(got) != 64 {
		t.Fatalf("source hash = %q, want 64 hex characters", got)
	}
	changed := doc
	changed.Source += "' changed without a version increment\n"
	if snapshot.Matches(changed) {
		t.Fatal("same-version source change matched the old snapshot")
	}
	wantLines := []string{"Option Explicit", "Public Sub First()", "  Dim value As Long", "End Sub", "Private Function Second() As Long", "End Function", ""}
	if got := snapshot.Lines(); !reflect.DeepEqual(got, wantLines) {
		t.Fatalf("lines = %#v, want %#v", got, wantLines)
	}
	lines := snapshot.Lines()
	lines[0] = "mutated"
	if snapshot.Lines()[0] != "Option Explicit" {
		t.Fatal("snapshot lines were mutated through returned slice")
	}
	procedures := snapshot.Procedures()
	if len(procedures) != 2 || procedures[0].Name != "First" || procedures[0].Range.Start.Line != 1 || procedures[0].Range.End.Line != 3 || procedures[1].Name != "Second" {
		t.Fatalf("procedures = %+v", procedures)
	}
	procedures[0].Name = "mutated"
	if snapshot.Procedures()[0].Name != "First" {
		t.Fatal("snapshot procedures were mutated through returned slice")
	}
	if name, scope := currentProcedureForDocument(view, Position{Line: 2}); name != "First" || scope == nil || scope.End.Line != 3 {
		t.Fatalf("procedure lookup = (%q, %+v)", name, scope)
	}
}

func TestNewIncrementalAnalysisSnapshotRejectsDivergedPreviousSource(t *testing.T) {
	previousDoc := Document{Path: "Main.bas", Source: "Sub A()\nEnd Sub\n", Version: 1}
	previous := NewAnalysisSnapshot(previousDoc)
	previous.parseDocument = func(string, []byte) (*vbaast.ParsedDocument, error) {
		return vbaast.ParseDocument("Main.bas", []byte("Sub Other()\nEnd Sub\n"))
	}
	_, err := NewIncrementalAnalysisSnapshot(
		Document{Path: "Main.bas", Source: "Sub B()\nEnd Sub\n", Version: 2},
		previous,
		[]tree_sitter.InputEdit{{StartByte: 4, OldEndByte: 5, NewEndByte: 5, StartPosition: tree_sitter.Point{Column: 4}, OldEndPosition: tree_sitter.Point{Column: 5}, NewEndPosition: tree_sitter.Point{Column: 5}}},
	)
	if !errors.Is(err, ErrIncrementalSnapshotUnavailable) {
		t.Fatalf("incremental snapshot error = %v, want ErrIncrementalSnapshotUnavailable", err)
	}
}

func TestAnalysisSnapshotReusesSemanticSourceMetadata(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Source: "Option Explicit\nDim customer As String\n", Version: 1})
	first := snapshot.identifiers()
	second := snapshot.identifiers()
	if len(first) != 3 || len(first[1]) != 4 {
		t.Fatalf("identifier metadata = %+v", first)
	}
	if &first[1][0] != &second[1][0] {
		t.Fatal("semantic identifier metadata was rebuilt for the same snapshot")
	}
}

func TestAnalysisSnapshotSourceSymbolsAreLazyConcurrentAndDefensive(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{URI: "file:///Main.bas", Path: "Main.bas", Source: "Sub Main()\nEnd Sub\n", Version: 1})
	var loads atomic.Int32
	load := func() ([]Symbol, error) {
		loads.Add(1)
		return []Symbol{{
			Name: "Main", Parameters: []Parameter{{Name: "value"}},
			Documentation: doccomments.SymbolDocumentation{Parameters: map[string]string{"value": "original"}},
		}}, nil
	}
	const readers = 24
	start := make(chan struct{})
	results := make(chan []Symbol, readers)
	var wg sync.WaitGroup
	for range readers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			<-start
			syms, _, err := snapshot.SourceSymbols(load)
			if err != nil {
				t.Error(err)
			}
			results <- syms
		}()
	}
	close(start)
	wg.Wait()
	close(results)
	for syms := range results {
		if len(syms) != 1 || syms[0].Name != "Main" {
			t.Fatalf("symbols = %+v", syms)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}
	syms, hit, err := snapshot.SourceSymbols(load)
	if err != nil || !hit {
		t.Fatalf("cached symbols = (hit=%v, err=%v)", hit, err)
	}
	syms[0].Parameters[0].Name = "mutated"
	syms[0].Documentation.Parameters["value"] = "mutated"
	again, _, _ := snapshot.SourceSymbols(load)
	if again[0].Parameters[0].Name != "value" || again[0].Documentation.Parameters["value"] != "original" {
		t.Fatalf("cached symbols were mutated: %+v", again[0])
	}
}

func TestAnalysisSnapshotCachesDeterministicSymbolErrorAndRetires(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Source: "broken", Version: 1})
	want := errors.New("parse failed")
	var loads atomic.Int32
	load := func() ([]Symbol, error) { loads.Add(1); return nil, want }
	for i := 0; i < 2; i++ {
		if _, hit, err := snapshot.SourceSymbols(load); !errors.Is(err, want) || hit != (i > 0) {
			t.Fatalf("call %d = (hit=%v, err=%v)", i, hit, err)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}
	snapshot.Retire()
	snapshot.Retire()
	if !snapshot.Retired() {
		t.Fatal("snapshot was not retired")
	}
}

func TestAnalysisSnapshotRawCallSitesAreLazyConcurrentAndDefensive(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{URI: "file:///Main.bas", Path: "Main.bas", Source: "Sub Main()\n  Target value:=1\nEnd Sub\n", Version: 1})
	receiver := "service"
	var loads atomic.Int32
	load := func() (calls.FileResult, error) {
		loads.Add(1)
		return calls.FileResult{
			Path:       "Main.bas",
			ModuleName: "Main",
			ModuleKind: "standard",
			CallSites: []calls.CallSite{{
				File:   "Main.bas",
				Module: "Main",
				Caller: &calls.Caller{Name: "Main", Kind: "sub", QualifiedName: "Main.Main"},
				Callee: calls.Callee{Text: "service.Target", BaseName: "Target", Receiver: &receiver, Member: "Target"},
				Arguments: calls.Arguments{
					Count: 1,
					Named: []calls.NamedArgument{{Name: "value", ValueText: "1"}},
				},
			}},
		}, nil
	}
	const readers = 24
	start := make(chan struct{})
	results := make(chan calls.FileResult, readers)
	var wg sync.WaitGroup
	for range readers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			<-start
			result, _, err := snapshot.RawCallSites(load)
			if err != nil {
				t.Error(err)
			}
			results <- result
		}()
	}
	close(start)
	wg.Wait()
	close(results)
	for result := range results {
		if len(result.CallSites) != 1 || result.CallSites[0].Callee.BaseName != "Target" {
			t.Fatalf("call sites = %+v", result.CallSites)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}

	result, hit, err := snapshot.RawCallSites(load)
	if err != nil || !hit {
		t.Fatalf("cached call sites = (hit=%v, err=%v)", hit, err)
	}
	result.CallSites[0].Caller.Name = "mutated"
	*result.CallSites[0].Callee.Receiver = "mutated"
	result.CallSites[0].Arguments.Named[0].Name = "mutated"
	again, _, _ := snapshot.RawCallSites(load)
	site := again.CallSites[0]
	if site.Caller.Name != "Main" || *site.Callee.Receiver != "service" || site.Arguments.Named[0].Name != "value" {
		t.Fatalf("cached call site was mutated: %+v", site)
	}
}

func TestAnalysisSnapshotCachesDeterministicRawCallSiteError(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Source: "broken", Version: 1})
	want := errors.New("call extraction failed")
	var loads atomic.Int32
	load := func() (calls.FileResult, error) {
		loads.Add(1)
		return calls.FileResult{Path: "Main.bas"}, want
	}
	for i := 0; i < 2; i++ {
		result, hit, err := snapshot.RawCallSites(load)
		if !errors.Is(err, want) || hit != (i > 0) || result.Path != "Main.bas" {
			t.Fatalf("call %d = (result=%+v, hit=%v, err=%v)", i, result, hit, err)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}
}

func TestAnalysisSnapshotProcedureIRIsLazyConcurrentAndDefensive(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{
		URI: "file:///Main.bas", Path: "Main.bas", Source: "Sub Main()\nEnd Sub\n", Version: 1,
	})
	var loads atomic.Int32
	load := func() (procedureir.DocumentIR, error) {
		loads.Add(1)
		return procedureir.DocumentIR{
			Path: "Main.bas",
			Procedures: []procedureir.ProcedureIR{{
				Symbol: procedureir.ProcedureSymbol{
					Name: "Main", Parameters: []procedureir.Parameter{{Name: "value"}},
				},
				Statements: []procedureir.Statement{{
					ID: 1, Kind: procedureir.StatementCall, Text: "Target",
				}},
			}},
		}, nil
	}
	const readers = 24
	start := make(chan struct{})
	results := make(chan procedureir.DocumentIR, readers)
	var wg sync.WaitGroup
	for range readers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			<-start
			result, _, err := snapshot.ProcedureIR(load)
			if err != nil {
				t.Error(err)
			}
			results <- result
		}()
	}
	close(start)
	wg.Wait()
	close(results)
	for result := range results {
		if len(result.Procedures) != 1 || result.Procedures[0].Symbol.Name != "Main" {
			t.Fatalf("procedure IR = %+v", result)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}

	result, hit, err := snapshot.ProcedureIR(load)
	if err != nil || !hit {
		t.Fatalf("cached procedure IR = (hit=%v, err=%v)", hit, err)
	}
	result.Procedures[0].Symbol.Name = "mutated"
	result.Procedures[0].Symbol.Parameters[0].Name = "mutated"
	result.Procedures[0].Statements[0].Text = "mutated"
	again, _, _ := snapshot.ProcedureIR(load)
	if again.Procedures[0].Symbol.Name != "Main" ||
		again.Procedures[0].Symbol.Parameters[0].Name != "value" ||
		again.Procedures[0].Statements[0].Text != "Target" {
		t.Fatalf("cached procedure IR was mutated: %+v", again)
	}
}

func TestAnalysisSnapshotCachesDeterministicProcedureIRError(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Source: "broken", Version: 1})
	want := errors.New("procedure IR build failed")
	var loads atomic.Int32
	load := func() (procedureir.DocumentIR, error) {
		loads.Add(1)
		return procedureir.DocumentIR{Path: "Main.bas"}, want
	}
	for i := 0; i < 2; i++ {
		result, hit, err := snapshot.ProcedureIR(load)
		if !errors.Is(err, want) || hit != (i > 0) || result.Path != "Main.bas" {
			t.Fatalf("call %d = (result=%+v, hit=%v, err=%v)", i, result, hit, err)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}
}

func TestAnalysisSnapshotControlFlowIsLazyConcurrentAndDefensive(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Path: "Main.bas", Source: "Sub Main()\nEnd Sub\n", Version: 1})
	var loads atomic.Int32
	load := func() (vbacfg.Document, error) {
		loads.Add(1)
		return vbacfg.Document{
			Path: "Main.bas",
			Graphs: []vbacfg.Graph{{
				Procedure: procedureir.ProcedureSymbol{Name: "Main"},
				Blocks:    []vbacfg.Block{{ID: 1, Kind: vbacfg.BlockEntry}},
				Entry:     1,
			}},
		}, nil
	}
	const readers = 24
	start := make(chan struct{})
	results := make(chan vbacfg.Document, readers)
	var wg sync.WaitGroup
	for range readers {
		wg.Add(1)
		go func() {
			defer wg.Done()
			<-start
			result, _, err := snapshot.ControlFlowGraphs(load)
			if err != nil {
				t.Error(err)
			}
			results <- result
		}()
	}
	close(start)
	wg.Wait()
	close(results)
	for result := range results {
		if len(result.Graphs) != 1 || result.Graphs[0].Procedure.Name != "Main" {
			t.Fatalf("control-flow document = %+v", result)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}

	result, hit, err := snapshot.ControlFlowGraphs(load)
	if err != nil || !hit {
		t.Fatalf("cached control flow = (hit=%v, err=%v)", hit, err)
	}
	result.Graphs[0].Procedure.Name = "mutated"
	result.Graphs[0].Blocks[0].Kind = vbacfg.BlockUnknownExit
	again, _, _ := snapshot.ControlFlowGraphs(load)
	if again.Graphs[0].Procedure.Name != "Main" || again.Graphs[0].Blocks[0].Kind != vbacfg.BlockEntry {
		t.Fatalf("cached control flow was mutated: %+v", again)
	}
}

func TestAnalysisSnapshotCachesDeterministicControlFlowError(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{Source: "broken", Version: 1})
	want := errors.New("control-flow build failed")
	var loads atomic.Int32
	load := func() (vbacfg.Document, error) {
		loads.Add(1)
		return vbacfg.Document{Path: "Main.bas"}, want
	}
	for i := 0; i < 2; i++ {
		result, hit, err := snapshot.ControlFlowGraphs(load)
		if !errors.Is(err, want) || hit != (i > 0) || result.Path != "Main.bas" {
			t.Fatalf("call %d = (result=%+v, hit=%v, err=%v)", i, result, hit, err)
		}
	}
	if loads.Load() != 1 {
		t.Fatalf("loads = %d, want 1", loads.Load())
	}
}

func TestAnalysisSnapshotSharesOneParsedDocumentAcrossDiagnosticsSymbolsAndSemanticTokens(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{
		URI:        "file:///Main.bas",
		Path:       "Main.bas",
		ModuleKind: "standard",
		Version:    1,
		Source:     "Attribute VB_Name = \"Main\"\nOption Explicit\nPublic Sub Run()\n  Dim found As Range\n  Set found = Range(\"A1\").Find(What:=\"x\")\n  Debug.Print found.Value\nEnd Sub\n",
	})
	var parses atomic.Int32
	snapshot.parseDocument = func(path string, source []byte) (*vbaast.ParsedDocument, error) {
		parses.Add(1)
		return vbaast.ParseDocument(path, source)
	}
	doc := snapshot.Document()
	analyzer := newTestAnalyzer(t)
	_ = analyzer.Diagnostics(doc)
	procedureIR, hit, err := snapshot.ProcedureIR(func() (procedureir.DocumentIR, error) {
		t.Fatal("diagnostics did not initialize the snapshot procedure IR")
		return procedureir.DocumentIR{}, nil
	})
	if err != nil || !hit || len(procedureIR.Procedures) != 1 {
		t.Fatalf("procedure IR = (result=%+v, hit=%v, err=%v)", procedureIR, hit, err)
	}
	controlFlow, hit, err := snapshot.ControlFlowGraphs(func() (vbacfg.Document, error) {
		t.Fatal("diagnostics did not initialize the snapshot control-flow graphs")
		return vbacfg.Document{}, nil
	})
	if err != nil || !hit || len(controlFlow.Graphs) != 1 {
		t.Fatalf("control flow = (result=%+v, hit=%v, err=%v)", controlFlow, hit, err)
	}
	if _, err := analyzer.DocumentSymbols(doc); err != nil {
		t.Fatal(err)
	}
	if _, err := analyzer.SemanticTokens(doc, []Document{doc}); err != nil {
		t.Fatal(err)
	}
	callResult, hit, err := snapshot.RawCallSites(func() (calls.FileResult, error) {
		return calls.ExtractIR(procedureIR), nil
	})
	if err != nil || hit {
		t.Fatalf("raw call sites = (result=%+v, hit=%v, err=%v)", callResult, hit, err)
	}
	if callResult.Parse != (symbols.ParseSummary{}) || len(callResult.CallSites) != 2 {
		t.Fatalf("raw call sites = %+v", callResult)
	}
	if parses.Load() != 1 || snapshot.ParseCount() != 1 {
		t.Fatalf("parsed documents = factory:%d snapshot:%d, want 1", parses.Load(), snapshot.ParseCount())
	}
}

func TestAnalysisSnapshotDocumentIndexIsLazyAndReused(t *testing.T) {
	snapshot := NewAnalysisSnapshot(Document{
		URI: "file:///Main.bas", Path: "Main.bas", Version: 7,
		Source: "Sub First()\n  Dim item As Object\n  Set item = New Collection\n  With item\n    .Count\n  End With\nEnd Sub\n",
	})
	var loads atomic.Int32
	load := func() ([]Symbol, error) {
		loads.Add(1)
		return []Symbol{{
			Name: "item", Kind: "local_variable", Parent: "First", ReturnType: "Object",
			Range: Range{Start: Position{Line: 1}, End: Position{Line: 1, Character: 17}},
		}}, nil
	}
	first, hit, err := snapshot.documentIndex(load)
	if err != nil || hit {
		t.Fatalf("first document index = (hit=%v, err=%v)", hit, err)
	}
	second, hit, err := snapshot.documentIndex(load)
	if err != nil || !hit || first != second {
		t.Fatalf("reused document index = (same=%v, hit=%v, err=%v)", first == second, hit, err)
	}
	if loads.Load() != 1 {
		t.Fatalf("symbol loads = %d, want 1", loads.Load())
	}
	if got := first.symbolsByName["item"]; len(got) != 1 || got[0].Parent != "First" {
		t.Fatalf("indexed declarations = %+v", got)
	}
	first.initAssignments()
	if got := first.assignmentsByName["item"]; len(got) != 1 || got[0].expression != "New Collection" || got[0].procedure != "First" {
		t.Fatalf("indexed assignments = %+v", got)
	}
	if block, ok := first.withBlockAt(Position{Line: 4}); !ok || first.withBlocks[block].receiver != "item" {
		t.Fatalf("indexed With block = (%d, %v, %+v)", block, ok, first.withBlocks)
	}
}
