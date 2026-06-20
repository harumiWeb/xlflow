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
    missingValue = 1
End Sub
`
	diagnostics = analyzer.Diagnostics(doc)
	if !hasDiagnostic(diagnostics, "VB029") {
		t.Fatalf("VB029 diagnostic missing: %+v", diagnostics)
	}

	doc.Source = `Option Explicit
Public Sub Run()
    Dim missingValue As Long
    missingValue = 1
End Sub
`
	if diagnostics := analyzer.Diagnostics(doc); len(diagnostics) != 0 {
		t.Fatalf("expected undeclared diagnostic to clear, got %+v", diagnostics)
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

func TestDiagnosticsConvertLintByteColumnsToUTF16AfterJapaneseText(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	line := `    Debug.Print "😀日本": Range("A1").Select`
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: "Option Explicit\nPublic Sub Run()\n" +
			line + "\n" +
			"End Sub\n",
	}

	diagnostics := analyzer.Diagnostics(doc)
	var selectDiag *Diagnostic
	for i := range diagnostics {
		if diagnostics[i].Code == "VB002" {
			selectDiag = &diagnostics[i]
			break
		}
	}
	if selectDiag == nil {
		t.Fatalf("VB002 diagnostic missing: %+v", diagnostics)
	}
	want := utf16Len(line[:strings.Index(line, "Select")])
	if selectDiag.Range.Start.Line != 2 || selectDiag.Range.Start.Character != want {
		t.Fatalf("VB002 range start = %+v, want line 2 character %d", selectDiag.Range.Start, want)
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

func TestDocumentSymbolsSkipEmptyNamesFromIncompleteSource(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     "Option Explicit\nPublic Sub \n    Dim \nEnd Sub\n",
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, sym := range symbols {
		if strings.TrimSpace(sym.Name) == "" {
			t.Fatalf("empty symbol name should not be returned: %+v", symbols)
		}
		if !rangeContains(sym.Range, sym.Selection) {
			t.Fatalf("symbol selection range must be contained in full range: %+v", sym)
		}
	}
}

func TestDocumentSymbolsKeepSelectionInsideRangeForBareIncompleteSource(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     "P",
	}

	symbols, err := analyzer.DocumentSymbols(doc)
	if err != nil {
		t.Fatal(err)
	}
	for _, sym := range symbols {
		if !rangeContains(sym.Range, sym.Selection) {
			t.Fatalf("symbol selection range must be contained in full range: %+v", sym)
		}
	}
}

func TestWorkspaceSymbolsPreferOpenDocumentOverFilesystemContent(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(moduleDir, "Main.bas")
	if err := os.WriteFile(path, []byte("Option Explicit\nPublic Sub OldName()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	open := Document{
		URI:    "file:///C:/work/src/modules/Main.bas",
		Path:   path,
		Source: "Option Explicit\nPublic Sub NewName()\nEnd Sub\n",
	}

	symbols, err := analyzer.WorkspaceSymbols([]Document{open}, "")
	if err != nil {
		t.Fatal(err)
	}
	if hasSymbol(symbols, "OldName") {
		t.Fatalf("stale filesystem symbol should be hidden while document is open: %+v", symbols)
	}
	if !hasSymbol(symbols, "NewName") {
		t.Fatalf("open document symbol missing: %+v", symbols)
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
	if got, ok := analyzer.ResolveExpressionType(`Range("A1").Font`); !ok || got != "Excel.Font" {
		t.Fatalf("Range(...).Font = %q, %v", got, ok)
	}
	if got, ok := analyzer.ResolveExpressionType(`Application.WorksheetFunction`); !ok || got != "Excel.WorksheetFunction" {
		t.Fatalf("Application.WorksheetFunction = %q, %v", got, ok)
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

func TestHoverAndCompletionsResolveBuiltInCollection(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private Sub Earlier(ByVal result As Boolean)
End Sub

Sub Test()
    Dim result As Collection
    Set result = New Collection
    result.Co
End Sub
`
	doc := Document{Path: filepath.Join(t.TempDir(), "Main.bas"), Source: source}

	hoverLine := `    Set result = New Collection`
	hover, err := analyzer.Hover(doc, Position{Line: 6, Character: utf16Len(hoverLine[:strings.Index(hoverLine, "result")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "VBA.Collection") {
		t.Fatalf("unexpected Collection hover: %+v", hover)
	}
	if strings.Contains(hover.Contents, "Boolean") {
		t.Fatalf("Collection hover should not use earlier Boolean declaration: %+v", hover)
	}

	completionLine := `    result.Co`
	items, err := analyzer.Completions(doc, Position{Line: 7, Character: utf16Len(completionLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Count") {
		t.Fatalf("Collection.Count completion missing: %+v", items)
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

	fontSource := "Option Explicit\nSub Test()\n    Range(\"A1\").Font.Co\nEnd Sub\n"
	fontLine := `    Range("A1").Font.Co`
	fontDoc := Document{Path: filepath.Join(t.TempDir(), "Formatting.bas"), Source: fontSource}
	items, err = analyzer.Completions(fontDoc, Position{Line: 2, Character: utf16Len(fontLine)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Color") {
		t.Fatalf("Font.Color completion missing: %+v", items)
	}

	globalDoc := Document{Path: filepath.Join(t.TempDir(), "Globals.bas"), Source: "Option Explicit\nSub Test()\n    xlU\nEnd Sub\n"}
	items, err = analyzer.Completions(globalDoc, Position{Line: 2, Character: utf16Len("    xlU")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlUp") {
		t.Fatalf("xlUp completion missing: %+v", items)
	}
	layoutDoc := Document{Path: filepath.Join(t.TempDir(), "Layout.bas"), Source: "Option Explicit\nSub Test()\n    xlLand\nEnd Sub\n"}
	items, err = analyzer.Completions(layoutDoc, Position{Line: 2, Character: utf16Len("    xlLand")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "xlLandscape") {
		t.Fatalf("xlLandscape completion missing: %+v", items)
	}
}

func TestCompletionsReturnTypeCandidatesInDeclarationContexts(t *testing.T) {
	root := t.TempDir()
	classDir := filepath.Join(root, "src", "classes")
	if err := os.MkdirAll(classDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(classDir, "MyService.cls"), []byte(`VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "MyService"
Option Explicit
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(root, "src", "modules", "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Dim ws As Wo\nEnd Sub\n",
	}
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim ws As Wo")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Workbook") {
		t.Fatalf("Workbook type alias completion missing: %+v", items)
	}
	if hasCompletion(items, "xlWorksheet") || hasCompletion(items, "ThisWorkbook") {
		t.Fatalf("type context should not include constants or globals: %+v", items)
	}
	workbook, ok := findCompletion(items, "Workbook")
	if !ok || workbook.ReplaceRange == nil || workbook.ReplaceRange.Start.Character != utf16Len("    Dim ws As ") {
		t.Fatalf("Workbook should replace only typed type prefix: %+v", workbook)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Dim value As Str\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim value As Str")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "String") {
		t.Fatalf("String built-in type completion missing: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Set dict = New Di\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Set dict = New Di")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Dictionary") {
		t.Fatalf("Dictionary alias completion missing after New: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    Dim service As My\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Dim service As My")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "MyService") {
		t.Fatalf("class module type completion missing: %+v", items)
	}
}

func TestCompletionsReturnModuleProcedureCandidates(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	utilsPath := filepath.Join(moduleDir, "Utils.bas")
	if err := os.WriteFile(utilsPath, []byte(`Attribute VB_Name = "Utils"
Option Explicit
Public Function BuildName() As String
End Function
Private Function HiddenName() As String
End Function
Sub RunDefault()
End Sub
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(moduleDir, "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    Utils.Bu\nEnd Sub\n",
	}
	line := `    Utils.Bu`
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(line)}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "BuildName") {
		t.Fatalf("Utils.BuildName completion missing: %+v", items)
	}
	if hasCompletion(items, "HiddenName") {
		t.Fatalf("private module member should not be completed: %+v", items)
	}

	openUtils := Document{
		URI:    "file:///C:/work/src/modules/Utils.bas",
		Path:   utilsPath,
		Source: "Attribute VB_Name = \"Utils\"\nOption Explicit\nPublic Function UnsavedName() As String\nEnd Function\n",
	}
	doc.Source = "Option Explicit\nSub Test()\n    Utils.Un\nEnd Sub\n"
	line = `    Utils.Un`
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len(line)}, []Document{openUtils})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "UnsavedName") {
		t.Fatalf("open document module member completion missing: %+v", items)
	}
	if hasCompletion(items, "BuildName") {
		t.Fatalf("stale filesystem module member should be hidden while open document exists: %+v", items)
	}
}

func TestCompletionsHidePrivateSymbolsFromOtherModules(t *testing.T) {
	root := t.TempDir()
	moduleDir := filepath.Join(root, "src", "modules")
	if err := os.MkdirAll(moduleDir, 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(moduleDir, "Utils.bas"), []byte(`Attribute VB_Name = "Utils"
Option Explicit
Private Const SECRET_VALUE As Long = 1
Public Const PUBLIC_VALUE As Long = 2
`), 0o644); err != nil {
		t.Fatal(err)
	}

	analyzer := newTestAnalyzer(t)
	analyzer.RootDir = root
	doc := Document{
		Path:   filepath.Join(moduleDir, "Main.bas"),
		Source: "Option Explicit\nSub Test()\n    SE\nEnd Sub\n",
	}
	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    SE")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hasCompletion(items, "SECRET_VALUE") {
		t.Fatalf("private const from another module should not be completed: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Test()\n    PU\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    PU")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "PUBLIC_VALUE") {
		t.Fatalf("public const from another module should be completed: %+v", items)
	}

	doc.Source = "Option Explicit\nPrivate Const SECRET_VALUE As Long = 1\nSub Test()\n    SE\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 3, Character: utf16Len("    SE")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "SECRET_VALUE") {
		t.Fatalf("private const from the same document should remain available: %+v", items)
	}
}

func TestCompletionsHideLocalsFromOtherProcedures(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit
Private Const MODULE_SECRET As Long = 1

Sub First()
    Dim firstLocal As Long
    fi
End Sub

Sub Second()
    Dim secondLocal As Long
    se
End Sub
`,
	}

	items, err := analyzer.Completions(doc, Position{Line: 5, Character: utf16Len("    fi")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "firstLocal") {
		t.Fatalf("current procedure local should be completed: %+v", items)
	}
	if hasCompletion(items, "secondLocal") {
		t.Fatalf("other procedure local should not be completed: %+v", items)
	}

	items, err = analyzer.Completions(doc, Position{Line: 10, Character: utf16Len("    se")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "secondLocal") {
		t.Fatalf("current procedure local should be completed: %+v", items)
	}
	if hasCompletion(items, "firstLocal") {
		t.Fatalf("other procedure local should not be completed: %+v", items)
	}

	doc.Source = strings.Replace(doc.Source, "    fi", "    MO", 1)
	items, err = analyzer.Completions(doc, Position{Line: 5, Character: utf16Len("    MO")}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "MODULE_SECRET") {
		t.Fatalf("same module private const should remain available: %+v", items)
	}
}

func TestCompletionsReturnModuleDeclarationSnippets(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := Document{
		Path: filepath.Join(t.TempDir(), "Main.bas"),
		Source: `Option Explicit

Pu
Public Sub Existing()
End Sub
`,
	}

	items, err := analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("Pu")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	publicSub, ok := findCompletion(items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing: %+v", items)
	}
	if !publicSub.Snippet || !strings.Contains(publicSub.InsertText, "End Sub") {
		t.Fatalf("Public Sub should be a procedure snippet: %+v", publicSub)
	}
	if !hasCompletion(items, "Public Function") {
		t.Fatalf("Public Function completion missing: %+v", items)
	}
	doc.Source = "Option Explicit\n\nPublic S\nPublic Sub Existing()\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("Public S")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	publicSub, ok = findCompletion(items, "Public Sub")
	if !ok {
		t.Fatalf("Public Sub completion missing for multi-word prefix: %+v", items)
	}
	if publicSub.ReplaceRange == nil || publicSub.ReplaceRange.Start.Character != 0 || publicSub.ReplaceRange.End.Character != utf16Len("Public S") {
		t.Fatalf("Public Sub should replace the typed declaration prefix: %+v", publicSub.ReplaceRange)
	}

	doc.Source = "Option Explicit\nSub Existing()\n    Pu\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Pu")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hasCompletion(items, "Public Sub") {
		t.Fatalf("module-level declaration snippets should not appear inside procedures: %+v", items)
	}

	doc.Source = "Option Explicit\nSub Existing()\n    Di\nEnd Sub\n"
	items, err = analyzer.Completions(doc, Position{Line: 2, Character: utf16Len("    Di")}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if !hasCompletion(items, "Dim") {
		t.Fatalf("procedure Dim snippet should appear inside procedures: %+v", items)
	}
	dimSnippet, ok := findCompletion(items, "Dim")
	if !ok || !dimSnippet.Snippet || !strings.Contains(dimSnippet.InsertText, "As ${2:Variant}") {
		t.Fatalf("Dim should be a local declaration snippet: %+v", dimSnippet)
	}
}

func TestHoverUsesConstantMetadata(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	line := "    Debug.Print xlLandscape"
	doc := Document{Path: filepath.Join(t.TempDir(), "Constants.bas"), Source: "Option Explicit\nSub Test()\n" + line + "\nEnd Sub\n"}

	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "xlLandscape")+3])}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "XlPageOrientation") {
		t.Fatalf("unexpected constant hover: %+v", hover)
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

func TestDefinitionAndReferencesPreferCurrentProcedureLocal(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First()
    Dim value As Long
    value = 1
End Sub

Sub Second()
    Dim value As Long
    value = 2
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 3, Character: utf16Len("    val")}

	defs, err := analyzer.Definition(doc, pos, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) != 1 || defs[0].Range.Start.Line != 2 {
		t.Fatalf("definition = %+v, want First.value declaration only", defs)
	}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 2 || withDecl[0].Range.Start.Line != 2 || withDecl[1].Range.Start.Line != 3 {
		t.Fatalf("local references with declaration = %+v, want First declaration and assignment only", withDecl)
	}

	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 1 || withoutDecl[0].Range.Start.Line != 3 {
		t.Fatalf("local references without declaration = %+v, want First assignment only", withoutDecl)
	}
}

func TestDefinitionAndReferencesResolveCurrentProcedureParameter(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First(ByVal value As Long)
    value = value + 1
End Sub

Sub Second(ByVal value As Long)
    value = value + 2
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	pos := Position{Line: 2, Character: utf16Len("    val")}

	defs, err := analyzer.Definition(doc, pos, []Document{doc}, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(defs) != 1 || defs[0].Range.Start.Line != 1 {
		t.Fatalf("parameter definition = %+v, want First.value parameter only", defs)
	}

	withDecl, err := analyzer.References(doc, pos, []Document{doc}, true, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withDecl) != 3 || withDecl[0].Range.Start.Line != 1 || withDecl[1].Range.Start.Line != 2 || withDecl[2].Range.Start.Line != 2 {
		t.Fatalf("parameter references with declaration = %+v, want First parameter and body references only", withDecl)
	}

	withoutDecl, err := analyzer.References(doc, pos, []Document{doc}, false, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(withoutDecl) != 2 || withoutDecl[0].Range.Start.Line != 2 || withoutDecl[1].Range.Start.Line != 2 {
		t.Fatalf("parameter references without declaration = %+v, want First body references only", withoutDecl)
	}
}

func TestHoverUsesCurrentProcedureParameterSymbol(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Sub First(value)
    Debug.Print value
End Sub

Sub Second(value As String)
    Debug.Print value
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
	line := `    Debug.Print value`
	hover, err := analyzer.Hover(doc, Position{Line: 2, Character: utf16Len(line[:strings.Index(line, "value")+3])}, []Document{doc})
	if err != nil {
		t.Fatal(err)
	}
	if hover == nil || !strings.Contains(hover.Contents, "value") {
		t.Fatalf("parameter hover missing: %+v", hover)
	}
	if strings.Contains(hover.Contents, "String") {
		t.Fatalf("parameter hover should not use another procedure parameter: %+v", hover)
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
	_, ok := findCompletion(items, label)
	return ok
}

func findCompletion(items []Completion, label string) (Completion, bool) {
	for _, item := range items {
		if item.Label == label {
			return item, true
		}
	}
	return Completion{}, false
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
