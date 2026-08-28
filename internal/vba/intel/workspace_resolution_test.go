package intel

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestWorkspaceResolutionViewQueriesIndexedSnapshot(t *testing.T) {
	symbols := []Symbol{
		{Name: "Run", Module: "Alpha", Kind: "sub", File: "a.bas", Parameters: []Parameter{{Name: "value", Type: "Long"}}},
		{Name: "Runner", Module: "Beta", Kind: "function", File: "b.bas"},
	}
	view := NewWorkspaceResolutionView(symbols)
	symbols[0].Name = "mutated input"
	symbols[0].Parameters[0].Name = "mutated input"
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
	got := view.Query(WorkspaceSymbolQuery{Text: "run", Mode: WorkspaceSymbolQueryExact})
	got[0].Name = "mutated"
	got[0].Parameters[0].Name = "mutated"
	got = view.Query(WorkspaceSymbolQuery{Text: "run", Mode: WorkspaceSymbolQueryExact})
	if len(got) != 1 || got[0].Name != "Run" || got[0].Parameters[0].Name != "value" {
		t.Fatalf("query result mutation changed immutable view: %+v", got)
	}
	duplicates := NewWorkspaceResolutionView([]Symbol{
		{Name: "Run", Module: "Second", Kind: "sub", File: "z.bas"},
		{Name: "Run", Module: "First", Kind: "sub", File: "a.bas"},
	})
	ordered := duplicates.Query(WorkspaceSymbolQuery{Text: "run", Mode: WorkspaceSymbolQueryExact})
	if len(ordered) != 2 || ordered[0].File != "a.bas" || ordered[1].File != "z.bas" {
		t.Fatalf("duplicate candidates are not deterministic: %+v", ordered)
	}
}

func BenchmarkWorkspaceResolutionViewLookup(b *testing.B) {
	for _, symbolCount := range []int{100, 1000, 10000} {
		b.Run(fmt.Sprintf("%d-symbols", symbolCount), func(b *testing.B) {
			symbols := make([]Symbol, 0, symbolCount)
			for i := 0; i < symbolCount; i++ {
				symbols = append(symbols, Symbol{
					Name:   fmt.Sprintf("Procedure%05d", i),
					Module: fmt.Sprintf("Module%03d", i%100),
					Kind:   "sub",
					File:   fmt.Sprintf("Module%03d.bas", i%100),
				})
			}
			view := NewWorkspaceResolutionView(symbols)
			b.ReportAllocs()
			b.ResetTimer()
			resultCount := 0
			for i := 0; i < b.N; i++ {
				resultCount += len(view.Query(WorkspaceSymbolQuery{Text: "procedure00042", Mode: WorkspaceSymbolQueryExact}))
				resultCount += len(view.Query(WorkspaceSymbolQuery{Text: "module042.procedure00042", Mode: WorkspaceSymbolQueryQualified}))
			}
			b.ReportMetric(float64(resultCount)/float64(b.N), "results/op")
		})
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
	if symbols, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{Text: "Run", Mode: WorkspaceSymbolQueryQualified}); !handled || len(symbols) != 1 || symbols[0].Name != "Run" {
		t.Fatalf("unqualified qualified query = (%+v, %t), want Run", symbols, handled)
	}
	aliased := Document{Path: "FileName.bas", Source: "Attribute VB_Name = \"PublicModule\"\nPublic Sub Run()\nEnd Sub\n", ModuleKind: "standard"}
	if symbols, handled := analyzer.LightweightDocumentSymbols(aliased, WorkspaceSymbolQuery{Text: "PublicModule.Run", Mode: WorkspaceSymbolQueryQualified}); !handled || len(symbols) != 1 || symbols[0].Name != "Run" {
		t.Fatalf("attribute-qualified query = (%+v, %t), want Run", symbols, handled)
	}
}

func TestLightweightDocumentSymbolsUsesSnapshotPrefixIndex(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure()\nEnd Sub\nPublic Function ReadValue() As Long\nEnd Function\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	doc.Snapshot = snapshot
	defer snapshot.Retire()
	analyzer := Analyzer{DB: vbadb.New(), DocumentSymbolsFunc: func(Document, DocumentSymbolLoader) ([]Symbol, error) {
		t.Fatal("prefix query unexpectedly built the full document symbol model")
		return nil, nil
	}}

	query := WorkspaceSymbolQuery{Text: "read", Mode: WorkspaceSymbolQueryPrefix, Interactive: true, DocumentPath: doc.Path}
	got, handled := analyzer.LightweightDocumentSymbols(doc, query)
	if !handled || len(got) != 1 || got[0].Name != "ReadValue" {
		t.Fatalf("prefix query = (%+v, %t), want ReadValue handled by snapshot index", got, handled)
	}
	if gotBuilds, gotHits := snapshot.InteractiveIndexStats(); gotBuilds != 1 || gotHits != 0 {
		t.Fatalf("first prefix stats = (%d, %d), want (1, 0)", gotBuilds, gotHits)
	}
	got, handled = analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{Text: "runprocedure", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path})
	if !handled || len(got) != 1 || got[0].Name != "RunProcedure" {
		t.Fatalf("exact query = (%+v, %t), want RunProcedure handled by snapshot index", got, handled)
	}
	if gotBuilds, gotHits := snapshot.InteractiveIndexStats(); gotBuilds != 1 || gotHits != 1 {
		t.Fatalf("warm query stats = (%d, %d), want (1, 1)", gotBuilds, gotHits)
	}
}

func TestDeclarationSymbolsContextUsesSnapshotInteractiveIndex(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure(ByVal value As Long)\n  Dim localValue As Long\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DocumentSymbolsFunc: func(Document, DocumentSymbolLoader) ([]Symbol, error) {
		t.Fatal("declaration query unexpectedly used the full document symbol callback")
		return nil, nil
	}}
	got, err := analyzer.DeclarationSymbolsContext(context.Background(), doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(got) == 0 {
		t.Fatal("declaration index returned no symbols")
	}
	for _, symbol := range got {
		if symbol.Name == "localValue" {
			t.Fatalf("declaration index walked procedure body: %+v", got)
		}
	}
	if snapshot.ParseCount() != 0 {
		t.Fatalf("declaration index created a full parsed document: parse count %d", snapshot.ParseCount())
	}
}

func TestLightweightDocumentSymbolsLoadsOnlyCurrentProcedureLocals(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure()\n  Dim localValue As Long\nEnd Sub\nPublic Sub OtherProcedure()\n  Dim otherValue As Long\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	doc.Snapshot = snapshot
	defer snapshot.Retire()
	analyzer := Analyzer{DB: vbadb.New()}
	got, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "local", Mode: WorkspaceSymbolQueryPrefix, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	})
	if !handled || len(got) != 1 || got[0].Name != "localValue" || got[0].Parent == "" {
		t.Fatalf("current-procedure locals = (%+v, %t), want localValue only", got, handled)
	}
	if other, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "othervalue", Mode: WorkspaceSymbolQueryPrefix, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	}); !handled || len(other) != 0 {
		t.Fatalf("other-procedure local leaked into current scope = (%+v, %t)", other, handled)
	}
	if builds, _ := snapshot.InteractiveIndexStats(); builds != 1 {
		t.Fatalf("interactive index builds = %d, want 1", builds)
	}
	if builds, reuses := snapshot.ProcedureCatalogStats(); builds != 1 || reuses == 0 {
		t.Fatalf("procedure catalog stats = (%d, %d), want one build and reuse", builds, reuses)
	}
}

func TestCompactInteractiveIndexAcceptsColonSeparatedBlockTerminators(t *testing.T) {
	source := "Public Sub ColonBlocks()\n" +
		"  For i = 1 To 3: Next i\n" +
		"  Do: Loop\n" +
		"  With Application: End With\n" +
		"End Sub\n"
	doc := Document{Path: "ColonBlocks.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()

	procedures, _ := procedureIndexForLines(documentLines(doc))
	if compactDeclarationRecovery(documentLines(doc), procedures) {
		t.Fatal("colon-separated block statements were classified as recovery")
	}

	symbols, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "ColonBlocks", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path,
	})
	if !handled || len(symbols) != 1 || symbols[0].Name != "ColonBlocks" {
		t.Fatalf("colon-separated procedure = (%+v, %t), want lightweight result", symbols, handled)
	}
}

func TestCompactInteractiveIndexFallsBackForProcedureBodySyntaxRecovery(t *testing.T) {
	tests := []struct {
		name string
		body string
	}{
		{name: "invalid punctuation", body: "    ????"},
		{name: "unterminated string", body: "    Debug.Print \"unterminated"},
		{name: "repeated declaration keyword", body: "    Dim x As Double, Dim i As Long"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			source := "Public Sub Broken()\n" + test.body + "\nEnd Sub\n"
			doc := Document{Path: "Broken.bas", Source: source, ModuleKind: "standard", Version: 1}
			snapshot := NewAnalysisSnapshot(doc)
			defer snapshot.Retire()
			doc = snapshot.Document()

			symbols, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
				Text: "Broken", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path,
			})
			if handled || len(symbols) != 0 {
				t.Fatalf("body recovery = (%+v, %t), want conservative fallback", symbols, handled)
			}
		})
	}
}

func TestProcedureCatalogReevaluatesAfterFullParse(t *testing.T) {
	doc := Document{
		Path:       "Catalog.bas",
		Source:     "Public Sub Run()\n    Dim value As Long\nEnd Sub\n",
		ModuleKind: "standard",
		Version:    1,
	}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New()}
	if symbols, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "Run", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path,
	}); !handled || len(symbols) != 1 {
		t.Fatalf("compact declaration lookup = (%+v, %t), want one symbol", symbols, handled)
	}
	if catalog := procedureCatalogForDocument(doc); catalog.ReuseSafe {
		t.Fatal("compact catalog unexpectedly certified procedure reuse")
	}
	if _, err := snapshot.ParsedDocument(); err != nil {
		t.Fatal(err)
	}
	if catalog := procedureCatalogForDocument(doc); !catalog.ReuseSafe {
		t.Fatal("full parse did not restore procedure catalog reuse safety")
	}
}

func TestLightweightDocumentSymbolsSelectsAccessorAtRequestPosition(t *testing.T) {
	source := "Public Property Get Value() As String\n" +
		"    Dim getLocal As String\n" +
		"End Property\n" +
		"Public Property Let Value(ByVal newValue As Long)\n" +
		"    Dim letLocal As Long\n" +
		"End Property\n"
	doc := Document{Path: "Accessor.bas", Source: source, ModuleKind: "class", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	requestPosition := Position{Line: 4, Character: 10}

	symbols, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "letLocal", Mode: WorkspaceSymbolQueryPrefix, Interactive: true,
		DocumentPath: doc.Path, Procedure: "Value", RequestPosition: &requestPosition,
	})
	if !handled || len(symbols) != 1 || symbols[0].Name != "letLocal" || symbols[0].Parent != "Value" {
		t.Fatalf("accessor locals at request position = (%+v, %t), want Let-local symbol", symbols, handled)
	}
}

func TestLightweightDocumentSymbolsRecognizesLabeledLocalDeclaration(t *testing.T) {
	source := "Public value As Long\n" +
		"Public Sub Run()\n" +
		"Start: Dim value As String\n" +
		"End Sub\n"
	doc := Document{Path: "LabeledLocal.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()

	symbols, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "value", Mode: WorkspaceSymbolQueryExact, Interactive: true,
		DocumentPath: doc.Path, Procedure: "Run",
	})
	if !handled || len(symbols) != 1 || symbols[0].Parent != "Run" || symbols[0].ReturnType != "String" {
		t.Fatalf("labeled local declaration = (%+v, %t), want procedure-local value", symbols, handled)
	}
}

func TestLightweightDocumentSymbolsPreservesMultilineLocalShadowing(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic localValue As Long\nPublic Sub RunProcedure()\n  Dim localValue As Long, _\n      anotherValue As Long\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New()}
	got, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "localValue", Mode: WorkspaceSymbolQueryExact, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	})
	if !handled || len(got) != 1 || !strings.EqualFold(got[0].Parent, "RunProcedure") {
		t.Fatalf("multiline local shadowing = (%+v, %t), want current-procedure local", got, handled)
	}
}

func TestLightweightDocumentSymbolsPreservesParameterShadowing(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic value As String\nPublic Sub RunProcedure(ByVal value As Long)\n  Debug.Print value\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	got, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "value", Mode: WorkspaceSymbolQueryExact, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	})
	if !handled || len(got) != 1 || got[0].Parent != "RunProcedure" || got[0].ReturnType != "Long" {
		t.Fatalf("parameter shadowing = (%+v, %t), want RunProcedure parameter", got, handled)
	}
	qualified, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "CurrentModule.value", Mode: WorkspaceSymbolQueryQualified, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	})
	if !handled || len(qualified) != 1 || qualified[0].Parent != "" || qualified[0].ReturnType != "String" {
		t.Fatalf("qualified module lookup = (%+v, %t), want module value", qualified, handled)
	}
}

func TestLightweightDocumentSymbolsPreservesMultilineProcedureHeader(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure( _\n    ByVal value As Long)\n  Debug.Print value\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New()}
	got, handled := analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "value", Mode: WorkspaceSymbolQueryExact, Interactive: true,
		DocumentPath: doc.Path, Procedure: "RunProcedure",
	})
	if !handled || len(got) != 1 || got[0].Kind != "parameter" || got[0].Parent != "RunProcedure" {
		t.Fatalf("multiline procedure parameter = (%+v, %t), want value parameter", got, handled)
	}
	if got := snapshot.ParseCount(); got != 0 {
		t.Fatalf("multiline declaration index parsed full document %d times", got)
	}
}

func TestDeclarationSymbolsContextIncludesClosedUserFormControls(t *testing.T) {
	source := "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm\n   Begin MSForms.TextBox txtName\n   End\nEnd\nAttribute VB_Name = \"CustomerForm\"\n"
	doc := Document{Path: "CustomerForm.frm", Source: source, ModuleKind: "form", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New()}
	syms, err := analyzer.DeclarationSymbolsContext(context.Background(), doc)
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, sym := range syms {
		if strings.EqualFold(sym.Name, "txtName") && strings.EqualFold(sym.Kind, "field") {
			found = true
			break
		}
	}
	if !found {
		t.Fatalf("closed UserForm control missing from declarations: %+v", syms)
	}
}

func TestLightweightDocumentSymbolsFallsBackForRecoveredProcedure(t *testing.T) {
	doc := Document{
		Path:       "Broken.bas",
		Source:     "Public Sub Broken()\n    If True Then\nEnd Sub\n",
		ModuleKind: "standard",
		Version:    1,
	}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc.Snapshot = snapshot
	if symbols, handled := (Analyzer{DB: vbadb.New()}).LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "Broken", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path,
	}); handled || len(symbols) != 0 {
		t.Fatalf("recovered procedure = (%+v, %t), want conservative fallback", symbols, handled)
	}
}

func TestCompactInteractiveIndexDoesNotCertifyProcedureReuse(t *testing.T) {
	source := "Public Sub Broken()\n  value =\nEnd Sub\nSub Later()\nEnd Sub\n"
	doc := Document{Path: "Broken.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New()}
	_, _ = analyzer.LightweightDocumentSymbols(doc, WorkspaceSymbolQuery{
		Text: "Broken", Mode: WorkspaceSymbolQueryExact, Interactive: true, DocumentPath: doc.Path,
	})
	if catalog := procedureCatalogForDocument(doc); catalog.ReuseSafe {
		t.Fatal("compact interactive index certified procedure reuse without full recovery state")
	}
	if got := snapshot.ParseCount(); got != 0 {
		t.Fatalf("compact recovery check parsed full document %d times", got)
	}
}

func TestInteractiveHoverTypedLocalDoesNotBuildFullDocumentSymbols(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure()\n  Dim localValue As Long\n  Debug.Print localValue\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	analyzer := Analyzer{DB: vbadb.New(), DocumentSymbolsFunc: func(Document, DocumentSymbolLoader) ([]Symbol, error) {
		t.Fatal("interactive hover unexpectedly built the full document symbol model")
		return nil, nil
	}}
	analyzer.WorkspaceSymbolQueryFunc = func(open []Document, query WorkspaceSymbolQuery) ([]Symbol, error) {
		var out []Symbol
		for _, candidate := range open {
			symbols, handled := analyzer.LightweightDocumentSymbols(candidate, query)
			if !handled {
				return nil, fmt.Errorf("interactive query was not handled")
			}
			out = append(out, symbols...)
		}
		return out, nil
	}
	line := "  Debug.Print localValue"
	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(line[:strings.Index(line, "localValue")+3])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Long") {
		t.Fatalf("typed local hover = %+v, want Long", hover)
	}
	if got := snapshot.ParseCount(); got != 0 {
		t.Fatalf("typed local hover parsed full document %d times", got)
	}
}

func TestInteractiveMemberHoverTypedLocalDoesNotBuildFullDocumentSymbols(t *testing.T) {
	source := "Attribute VB_Name = \"CurrentModule\"\nPublic Sub RunProcedure()\n  Dim ws As Worksheet\n  ws.Range(\"A1\")\nEnd Sub\n"
	doc := Document{Path: "CurrentModule.bas", Source: source, ModuleKind: "standard", Version: 1}
	snapshot := NewAnalysisSnapshot(doc)
	defer snapshot.Retire()
	doc = snapshot.Document()
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	analyzer := Analyzer{DB: db, DocumentSymbolsFunc: func(Document, DocumentSymbolLoader) ([]Symbol, error) {
		t.Fatal("interactive member hover unexpectedly built the full document symbol model")
		return nil, nil
	}}
	line := "  ws.Range(\"A1\")"
	hover, err := analyzer.Hover(doc, Position{Line: 3, Character: utf16Len(line[:strings.Index(line, "Range")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Worksheet.Range") {
		t.Fatalf("typed local member hover = %+v, want Worksheet.Range", hover)
	}
	if got := snapshot.ParseCount(); got != 0 {
		t.Fatalf("typed local member hover parsed full document %d times", got)
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

func TestDiagnosticsResolvesWorkspacePublicVB029Declarations(t *testing.T) {
	doc := Document{
		Path:       "Form1.frm",
		Source:     "Option Explicit\nPublic Sub Run()\n  sharedFlag = True\n  hiddenFlag = False\n  missingValue = 1\nEnd Sub\n",
		ModuleKind: "form",
		Version:    1,
	}
	analyzer := Analyzer{
		DB: vbadb.New(),
		WorkspaceSymbolsSnapshotFunc: func([]Document) ([]Symbol, error) {
			return []Symbol{
				{
					Name:       "sharedFlag",
					Kind:       "module_variable",
					Module:     "Globals",
					ModuleKind: "standard",
					Visibility: "Public",
					File:       "Globals.bas",
				},
				{
					Name:       "hiddenFlag",
					Kind:       "module_variable",
					Module:     "Globals",
					ModuleKind: "standard",
					File:       "Globals.bas",
				},
			}, nil
		},
	}

	got := diagnosticsByCode(analyzer.Diagnostics(doc), "VB029")
	if len(got) != 2 || !strings.Contains(got[0].Message, "hiddenFlag") || !strings.Contains(got[1].Message, "missingValue") {
		t.Fatalf("VB029 diagnostics = %+v, want hiddenFlag and missingValue", got)
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

func TestFastDiagnosticsRecomputesPropertyAccessorContracts(t *testing.T) {
	oldSource := "Property Get Name() As String\nEnd Property\nProperty Let Name(ByVal value As String)\nEnd Property\n"
	newSource := "Property Get Name() As String\nEnd Property\nProperty Let Name(ByVal value As Long)\nEnd Property\n"
	oldDoc := Document{Path: "Module1.bas", Source: oldSource, ModuleKind: "standard", Version: 1}
	oldCatalog := procedureCatalogForDocument(oldDoc)
	cache := buildDiagnosticCache(oldCatalog, nil)
	newDoc := Document{Path: oldDoc.Path, Source: newSource, ModuleKind: oldDoc.ModuleKind, Version: 2}
	result := Analyzer{DB: vbadb.New()}.DiagnosticsRequestContext(context.Background(), DiagnosticRequest{
		Document: newDoc, Mode: DiagnosticModeFast, PreviousCache: cache,
		Changes: ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: 2}, End: Position{Line: 2}}}},
	})
	if len(diagnosticsByCode(result.Diagnostics, "VB049")) == 0 {
		t.Fatalf("property contract change was not recomputed in fast mode: %+v", result.Diagnostics)
	}
	if result.Cache == nil {
		t.Fatal("property fallback did not publish a refreshed diagnostic cache")
	}
	if result.Cache == cache {
		t.Fatal("property fallback reused the stale diagnostic cache")
	}
	if len(result.Cache.Catalog.Entries) != len(cache.Catalog.Entries) || result.Cache.Catalog.Entries[1].SignatureHash == cache.Catalog.Entries[1].SignatureHash {
		t.Fatalf("property fallback cache catalog was not refreshed: old=%#v new=%#v", cache.Catalog, result.Cache.Catalog)
	}

	validDoc := Document{Path: oldDoc.Path, Source: oldSource, ModuleKind: oldDoc.ModuleKind, Version: 3}
	validResult := Analyzer{DB: vbadb.New()}.DiagnosticsRequestContext(context.Background(), DiagnosticRequest{
		Document: validDoc, Mode: DiagnosticModeFast, PreviousCache: result.Cache,
		Changes: ProcedureChangeSet{Ranges: []Range{{Start: Position{Line: 2}, End: Position{Line: 2}}}},
	})
	if got := diagnosticsByCode(validResult.Diagnostics, "VB049"); len(got) != 0 {
		t.Fatalf("stale VB049 diagnostics survived invalid-to-valid edit: %+v", got)
	}
	if validResult.Cache == nil || validResult.Cache == result.Cache {
		t.Fatal("invalid-to-valid property edit did not publish a fresh cache")
	}
}
