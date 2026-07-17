package intel

import (
	"errors"
	"reflect"
	"sync"
	"sync/atomic"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vba/doccomments"
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
