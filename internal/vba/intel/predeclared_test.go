package intel

import (
	"fmt"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestPredeclaredClassDeclarationIndexStaysLazy(t *testing.T) {
	for _, procedures := range []int{0, 256} {
		t.Run(fmt.Sprint(procedures), func(t *testing.T) {
			var source strings.Builder
			source.WriteString("Attribute VB_Name = \"Predeclared\"\nAttribute VB_PredeclaredId = True\nOption Explicit\n")
			for i := range procedures {
				fmt.Fprintf(&source, "Public Sub Unrelated%d()\n  Dim hiddenLocal As Long\n  hiddenLocal = 1\nEnd Sub\n", i)
			}
			snapshot := NewAnalysisSnapshot(Document{Path: "Predeclared.cls", Source: source.String(), ModuleKind: "class", Version: 1})
			defer snapshot.Retire()
			analyzer := Analyzer{DocumentSymbolsFunc: func(Document, DocumentSymbolLoader) ([]Symbol, error) {
				t.Fatal("declaration lookup requested full symbols")
				return nil, nil
			}}
			for range 3 {
				symbols, err := analyzer.DeclarationSymbolsContext(t.Context(), snapshot.Document())
				if err != nil {
					t.Fatal(err)
				}
				if visible := visibleAssignmentDeclarations(symbols); len(visible) != 1 || !visible["Predeclared"] {
					t.Fatalf("assignment declarations = %+v, want only Predeclared", visible)
				}
			}
			if builds, hits := snapshot.InteractiveIndexStats(); builds != 1 || hits != 2 {
				t.Fatalf("index builds/hits = %d/%d, want 1/2", builds, hits)
			}
			if snapshot.ParseCount() != 0 || snapshot.procedureIRBuildCount.Load() != 0 || snapshot.cfgBuildCount.Load() != 0 {
				t.Fatal("predeclared metadata lookup built full parse, procedure IR or CFG")
			}
		})
	}
}

func TestDiagnosticsVB029PredeclaredClassSnapshot(t *testing.T) {
	for _, attribute := range []string{"True", "False", ""} {
		t.Run(attribute, func(t *testing.T) {
			analyzer := Analyzer{DB: vbadb.New(), Config: config.Default()}
			classSource := "Attribute VB_Name = \"Predeclared\"\nOption Explicit\nPublic Value As Long\n"
			if attribute != "" {
				classSource = "Attribute VB_PredeclaredId = " + attribute + "\n" + classSource
			}
			classDoc := Document{Path: "DifferentFilename.cls", Source: classSource, ModuleKind: "class", Version: 2}
			classSymbols, err := analyzer.DeclarationSymbolsContext(t.Context(), classDoc)
			if err != nil {
				t.Fatal(err)
			}
			// Project declarations come from the same snapshot as the unsaved class.
			analyzer.WorkspaceSymbolsSnapshotFunc = func([]Document) ([]Symbol, error) { return classSymbols, nil }
			doc := Document{Path: "UsePredeclared.bas", ModuleKind: "standard", Version: 1,
				Source: "Attribute VB_Name = \"UsePredeclared\"\nOption Explicit\nPublic Sub Run()\n  Set pReDeClArEd = Nothing\n  Set missingValue = Nothing\nEnd Sub\n"}
			got := diagnosticsByCode(analyzer.Diagnostics(doc), "VB029")
			want := 2
			if attribute == "True" {
				want = 1
			}
			if len(got) != want {
				t.Fatalf("VB029 = %+v, want %d", got, want)
			}
			if got[len(got)-1].Range.Start.Line != 4 {
				t.Fatalf("missingValue diagnostic = %+v", got)
			}
		})
	}
}
