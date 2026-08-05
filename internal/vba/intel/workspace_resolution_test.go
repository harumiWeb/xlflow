package intel

import (
	"context"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestWorkspaceResolutionViewQueriesIndexedSnapshot(t *testing.T) {
	view := NewWorkspaceResolutionView([]Symbol{
		{Name: "Run", Module: "Alpha", Kind: "sub", File: "a.bas"},
		{Name: "Runner", Module: "Beta", Kind: "function", File: "b.bas"},
	})
	tests := []struct {
		query WorkspaceSymbolQuery
		want  string
	}{
		{WorkspaceSymbolQuery{Text: "run", Mode: WorkspaceSymbolQueryExact}, "Run"},
		{WorkspaceSymbolQuery{Text: "alpha.run", Mode: WorkspaceSymbolQueryQualified}, "Run"},
		{WorkspaceSymbolQuery{Text: "beta", Mode: WorkspaceSymbolQueryModule}, "Runner"},
		{WorkspaceSymbolQuery{Text: "function", Mode: WorkspaceSymbolQueryKind}, "Runner"},
	}
	for _, test := range tests {
		got := view.Query(test.query)
		if len(got) != 1 || got[0].Name != test.want {
			t.Fatalf("Query(%+v) = %+v, want %q", test.query, got, test.want)
		}
	}
	if got := view.Query(WorkspaceSymbolQuery{Text: "run", Mode: WorkspaceSymbolQueryPrefix}); len(got) != 2 {
		t.Fatalf("prefix query returned %d symbols, want 2", len(got))
	}
}

func TestLightweightDocumentSymbolsHonorsQualifiedModule(t *testing.T) {
	doc := Document{Path: "CurrentModule.bas", Source: "Public Sub Run()\nEnd Sub\n", ModuleKind: "standard"}
	analyzer := Analyzer{DB: vbadb.New()}
	if symbols, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{Text: "OtherModule.Run", Mode: WorkspaceSymbolQueryQualified}); !handled || len(symbols) != 0 {
		t.Fatalf("foreign qualified query = (%+v, %t), want handled without current-module symbols", symbols, handled)
	}
	if symbols, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{Text: "CurrentModule.Run", Mode: WorkspaceSymbolQueryQualified}); !handled || len(symbols) != 1 || symbols[0].Name != "Run" {
		t.Fatalf("current qualified query = (%+v, %t), want Run", symbols, handled)
	}
}

func TestFastDiagnosticsDoNotReportMissingModulePreambleFromProcedureFragment(t *testing.T) {
	source := "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Sub Run()\n  Debug.Print 1\nEnd Sub\n"
	doc := Document{Path: "Module1.bas", Source: source, ModuleKind: "standard", Version: 2}
	result := Analyzer{DB: vbadb.New()}.DiagnosticsRequestContext(context.Background(), DiagnosticRequest{
		Document: doc,
		Mode:     DiagnosticModeFast,
		Changes:  ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: 3}, End: Position{Line: 3}}}},
	})
	for _, diagnostic := range result.Diagnostics {
		if diagnostic.Code == "VB001" || diagnostic.Code == "VB031" {
			t.Fatalf("Fast procedure fragment leaked module-level %s: %+v", diagnostic.Code, result.Diagnostics)
		}
	}
}

func TestFastDiagnosticsPreserveConditionalContextForLaterProcedure(t *testing.T) {
	source := "Attribute VB_Name = \"Module1\"\nOption Explicit\nSub Before()\nEnd Sub\n#If VBA7 Then\nSub Changed()\n  Dim unused As Long\nEnd Sub\n#End If\n"
	doc := Document{Path: "Module1.bas", Source: source, ModuleKind: "standard", Version: 2}
	catalog := procedureCatalogForDocument(doc)
	if len(catalog.Entries) != 2 {
		t.Fatalf("procedure catalog entries = %d, want 2", len(catalog.Entries))
	}
	entry := catalog.Entries[1]
	preambleEnd := catalog.Entries[0].StartByte
	fragment := fastDiagnosticFragmentSource(source, source[:preambleEnd], preambleEnd, entry)
	if !strings.Contains(fragment, "#If VBA7 Then\nSub Changed()") || strings.Count(fragment, "#End If") != 1 {
		t.Fatalf("conditional fragment = %q, want enclosing #If/#End If", fragment)
	}
	fragmentCatalog := procedureCatalogForDocumentMode(Document{Path: doc.Path, Source: fragment, ModuleKind: doc.ModuleKind}, false)
	if len(fragmentCatalog.Entries) != 1 || fragmentCatalog.Entries[0].Identity.CanonicalName != "changed" {
		t.Fatalf("fragment procedure catalog = %+v, want only Changed", fragmentCatalog.Entries)
	}
}

func TestDiagnosticsBuildsOneWorkspaceResolutionSnapshot(t *testing.T) {
	source := "Sub Caller()\n  Target\nEnd Sub\n\nSub Target()\nEnd Sub\n"
	doc := Document{Path: "Module1.bas", URI: "file:///Module1.bas", Source: source, ModuleKind: "standard", Version: 1}
	doc.Snapshot = NewAnalysisSnapshot(doc)
	defer doc.Snapshot.Retire()

	snapshots := 0
	queries := 0
	analyzer := Analyzer{
		DB: vbadb.New(),
		WorkspaceSymbolsSnapshotFunc: func(open []Document) ([]Symbol, error) {
			snapshots++
			return []Symbol{{Name: "Target", Module: "Module1", Kind: "sub", File: doc.Path}}, nil
		},
		WorkspaceSymbolQueryFunc: func([]Document, WorkspaceSymbolQuery) ([]Symbol, error) {
			queries++
			return nil, nil
		},
	}
	_ = analyzer.DiagnosticsContext(context.Background(), doc)
	if snapshots != 1 {
		t.Fatalf("workspace snapshots = %d, want 1", snapshots)
	}
	if queries != 0 {
		t.Fatalf("fallback workspace queries = %d, want 0", queries)
	}
}

func TestAnalysisSnapshotSourceMapHandlesCRLFAndUTF16(t *testing.T) {
	doc := Document{Source: "A\r\n😀B\rC", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	doc.Snapshot = snapshot
	defer snapshot.Retire()

	if got := byteOffsetForDocumentPosition(doc, Position{Line: 1, Character: 2}); got != 7 {
		t.Fatalf("byte offset = %d, want 7", got)
	}
	if got := positionForDocumentByteOffset(doc, 7); got != (Position{Line: 1, Character: 2}) {
		t.Fatalf("position = %+v, want line 1 character 2", got)
	}
	if got := byteOffsetForDocumentPosition(doc, Position{Line: 2, Character: 1}); got != len(doc.Source) {
		t.Fatalf("final-line byte offset = %d, want %d", got, len(doc.Source))
	}
}

func TestFastDiagnosticsRebasesUnchangedProcedureDiagnostics(t *testing.T) {
	oldSource := "Sub A()\n  x = 1\nEnd Sub\nSub B()\n  Dim unused As Long\nEnd Sub\n"
	oldDoc := Document{Path: "Module1.bas", Source: oldSource, ModuleKind: "standard", Version: 1}
	oldCatalog := procedureCatalogForDocument(oldDoc)
	oldB := oldCatalog.Entries[1]
	cachedDiagnostic := Diagnostic{Code: "VB999", Severity: "warning", Range: Range{
		Start: Position{Line: oldB.Range.Start.Line + 1, Character: 2},
		End:   Position{Line: oldB.Range.Start.Line + 1, Character: 5},
	}}
	cache := buildDiagnosticCache(oldCatalog, []Diagnostic{cachedDiagnostic})

	newSource := "Sub A()\n  x = 1\n\nEnd Sub\nSub B()\n  Dim unused As Long\nEnd Sub\n"
	newDoc := Document{Path: oldDoc.Path, Source: newSource, ModuleKind: oldDoc.ModuleKind, Version: 2}
	result := Analyzer{DB: vbadb.New()}.DiagnosticsRequestContext(context.Background(), DiagnosticRequest{
		Document: newDoc, Mode: DiagnosticModeFast, PreviousCache: cache,
		Changes: ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: 2}, End: Position{Line: 2}}}},
	})
	for _, diagnostic := range result.Diagnostics {
		if diagnostic.Code == "VB999" {
			if diagnostic.Range.Start.Line != cachedDiagnostic.Range.Start.Line+1 {
				t.Fatalf("rebased line = %d, want %d", diagnostic.Range.Start.Line, cachedDiagnostic.Range.Start.Line+1)
			}
			return
		}
	}
	t.Fatalf("unchanged procedure diagnostic was not preserved: %+v", result.Diagnostics)
}

func TestFastDiagnosticsFallsBackForBroadInvalidation(t *testing.T) {
	tests := []struct {
		name       string
		oldSource  string
		newSource  string
		changeLine int
		want       int
	}{
		{
			name:       "module declaration",
			oldSource:  "Public sharedValue As Long\nSub A()\nEnd Sub\nSub B()\nEnd Sub\n",
			newSource:  "Public sharedValue As String\nSub A()\nEnd Sub\nSub B()\nEnd Sub\n",
			changeLine: 0,
			want:       2,
		},
		{
			name:       "conditional compilation inside procedure",
			oldSource:  "Sub A()\n#If VBA7 Then\n#End If\nEnd Sub\nSub B()\nEnd Sub\n",
			newSource:  "Sub A()\n#If Win64 Then\n#End If\nEnd Sub\nSub B()\nEnd Sub\n",
			changeLine: 1,
			want:       2,
		},
		{
			name:       "procedure signature",
			oldSource:  "Sub A()\nEnd Sub\nSub B()\nEnd Sub\n",
			newSource:  "Sub A(ByVal value As Long)\nEnd Sub\nSub B()\nEnd Sub\n",
			changeLine: 0,
			want:       1,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			oldCatalog := procedureCatalogForDocument(Document{Path: "Module1.bas", Source: test.oldSource, ModuleKind: "standard"})
			cache := buildDiagnosticCache(oldCatalog, nil)
			newCatalog := procedureCatalogForDocument(Document{Path: "Module1.bas", Source: test.newSource, ModuleKind: "standard"})
			changed := changedProcedureEntries(newCatalog, cache, ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: test.changeLine}, End: Position{Line: test.changeLine}}}})
			if len(changed) != test.want {
				t.Fatalf("changed procedures = %d, want %d: %+v", len(changed), test.want, changed)
			}
		})
	}
}
