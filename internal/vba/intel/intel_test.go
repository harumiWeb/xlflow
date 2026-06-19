package intel

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	"github.com/harumiWeb/xlflow/internal/vbadb"
)

func TestDiagnosticsHandlesMalformedSourceAndJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		URI:    "file:///C:/work/%E6%97%A5%E6%9C%AC/Main.bas",
		Path:   filepath.Join(t.TempDir(), "日本", "Main.bas"),
		Source: "Sub Broken()\n    Debug.Print \"日本語\"\n",
	}

	diagnostics := analyzer.Diagnostics(doc)
	if len(diagnostics) == 0 {
		t.Fatal("expected diagnostics for malformed source")
	}
	for _, diag := range diagnostics {
		if diag.Source != "xlflow" {
			t.Fatalf("diagnostic source = %q, want xlflow", diag.Source)
		}
		if diag.Range.Start.Character < 0 || diag.Range.End.Character < 0 {
			t.Fatalf("negative range in diagnostic: %+v", diag)
		}
	}
}

func TestDocumentSymbolsUseUnsavedDocumentContent(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source: `Attribute VB_Name = "Main"
Option Explicit
Public Sub RunReport()
End Sub
`,
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	if !hasSymbol(symbols, "RunReport") {
		t.Fatalf("RunReport not found in symbols: %+v", symbols)
	}
}

func TestResolveExpressionTypeHandlesExcelCollectionsAndCreateObject(t *testing.T) {
	analyzer := newTestAnalyzer(t)

	if got, ok := analyzer.ResolveExpressionType(`Worksheets("Input").Range("A1")`); !ok || got != "Excel.Range" {
		t.Fatalf("Worksheets(...).Range(...) = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Workbooks.Open("book.xlsx")`); !ok || got != "Excel.Workbook" {
		t.Fatalf("Workbooks.Open(...) = %q, %v", got, ok)
	}

	doc := Document{
		Source: `Option Explicit
Sub Test()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "a", 1
End Sub
`,
	}
	if got, ok := analyzer.inferWordType(doc, "dict"); !ok || got != "Scripting.Dictionary" {
		t.Fatalf("dict type = %q, %v", got, ok)
	}
}

func TestHoverUsesUTF16PositionsAfterJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Test()\n    Debug.Print \"日本語\" & Range(\"A1\").Value\nEnd Sub\n"
	line := `    Debug.Print "日本語" & Range("A1").Value`
	character := utf16Len(line[:strings.Index(line, "Range")+len("Ra")])
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: character}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "Excel.Range") {
		t.Fatalf("unexpected hover: %+v", hover)
	}
}

func newTestAnalyzer(t *testing.T) Analyzer {
	t.Helper()
	db, err := vbadb.LoadBuiltin()
	if err != nil {
		t.Fatal(err)
	}
	return Analyzer{RootDir: t.TempDir(), Config: config.Default(), DB: db}
}

func hasSymbol(symbols []Symbol, name string) bool {
	for _, symbol := range symbols {
		if symbol.Name == name {
			return true
		}
	}
	return false
}
