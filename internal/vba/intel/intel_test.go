package intel

import (
	"os"
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

func TestDiagnosticsUseSharedLintRulesForUnsavedContent(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Public Sub Run()
    Range("A1").Select
End Sub
`,
	}

	diagnostics := analyzer.Diagnostics(doc)
	if !hasDiagnostic(diagnostics, "VB002") {
		t.Fatalf("VB002 diagnostic missing: %+v", diagnostics)
	}

	doc.Source = `Option Explicit
Public Sub Run()
    Range("A1").Value = 1
End Sub
`
	if diagnostics := analyzer.Diagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("expected diagnostics to clear, got %+v", diagnostics)
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

func TestCompletionsReturnMemberAndGlobalCandidates(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := "Option Explicit\nSub Test()\n    Worksheets(\"Input\").Ra\nEnd Sub\n"
	memberLine := `    Worksheets("Input").Ra`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(memberLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Range") {
		t.Fatalf("Range completion missing: %+v", items)
	}

	globalDoc := Document{Path: filepath.Join(t.TempDir(), "Globals.bas"), Source: "Option Explicit\nSub Test()\n    xlU\nEnd Sub\n"}
	items, err = analyzer.Completions(globalDoc, Position{Line: 2, Character: utf16Len("    xlU")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlUp") {
		t.Fatalf("xlUp completion missing: %+v", items)
	}
}

func TestUserFormControlsResolveForHoverCompletionAndDefinition(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Begin MSForms.TextBox txtName
   End
End
Attribute VB_Name = "CustomerForm"
Option Explicit
Private Sub UserForm_Initialize()
    Me.txtName.Text = "Alice"
    Me.Controls("txtName").Text = "Bob"
End Sub
`
	doc := Document{
		URI:        "file:///C:/work/src/forms/CustomerForm.frm",
		Path:       filepath.Join(t.TempDir(), "CustomerForm.frm"),
		ModuleKind: "form",
		Source:     source,
	}

	hoverLine := `    Me.txtName.Text = "Alice"`
	hover, err := analyzer.Hover(doc, Position{Line: 8, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "txtName")+3])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "MSForms.TextBox") {
		t.Fatalf("unexpected hover: %+v", hover)
	}

	controlLine := `    Me.Controls("txtName").Te`
	items, err := analyzer.Completions(doc, Position{Line: 9, Character: utf16Len(controlLine)}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Text") {
		t.Fatalf("Text completion missing: %+v", items)
	}

	defs, err := analyzer.Definition(doc, Position{Line: 8, Character: utf16Len(`    Me.txtN`)}, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) == 0 || defs[0].Range.Start.Line != 2 {
		t.Fatalf("control definition = %+v, want declaration on .frm Begin line", defs)
	}
}

func TestUserFormControlsResolveFromSidecarCodeBehind(t *testing.T) {
	root := t.TempDir()
	formsDir := filepath.Join(root, "src", "forms")
	codeDir := filepath.Join(formsDir, "code")
	if err := os.MkdirAll(codeDir, 0o755); err != nil {
		t.Fatal(err)
	}
	frm := `VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerForm
   Begin MSForms.TextBox txtName
   End
End
Attribute VB_Name = "CustomerForm"
`
	if err := os.WriteFile(filepath.Join(formsDir, "CustomerForm.frm"), []byte(frm), 0o644); err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:       filepath.Join(codeDir, "CustomerForm.bas"),
		ModuleKind: "form",
		Source:     "Option Explicit\nPrivate Sub UserForm_Initialize()\n    Me.txtName.Text = \"Alice\"\nEnd Sub\n",
	}
	line := `    Me.txtName.Text = "Alice"`
	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "txtName")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "MSForms.TextBox") {
		t.Fatalf("unexpected sidecar hover: %+v", hover)
	}
}

func TestReferencesFindOpenDocumentOccurrencesAndCanSkipDeclaration(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Public Sub RunReport()
    RunReport
End Sub
`
	doc := Document{
		URI:        "file:///C:/work/src/modules/Main.bas",
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 1, Character: utf16Len("Public Sub Run")}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 2 {
		t.Fatalf("references with declaration = %d, want 2: %+v", len(withDecl), withDecl)
	}
	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 1 || withoutDecl[0].Range.Start.Line != 2 {
		t.Fatalf("references without declaration = %+v, want only call line", withoutDecl)
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

func hasCompletion(items []Completion, label string) bool {
	for _, item := range items {
		if item.Label == label {
			return true
		}
	}
	return false
}

func hasSymbol(symbols []Symbol, name string) bool {
	for _, symbol := range symbols {
		if symbol.Name == name {
			return true
		}
	}
	return false
}

func hasDiagnostic(diagnostics []Diagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}
