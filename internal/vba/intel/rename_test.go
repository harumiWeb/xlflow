package intel

import (
	"path/filepath"
	"sort"
	"strings"
	"testing"
)

func TestRenameLocalVariableParameterAndConst(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Public Sub Sample(ByVal value As String)
    Dim lastRow As Long
    Const LocalLimit As Long = 3
    lastRow = LocalLimit
    Debug.Print value, lastRow, "lastRow", LocalLimit ' lastRow
End Sub
`
	doc := renameDoc(t, source)

	renamed := mustRenameSource(t, analyzer, doc, "lastRow = LocalLimit", "lastRow", "lastDataRow")
	if !strings.Contains(renamed, "Dim lastDataRow As Long") ||
		!strings.Contains(renamed, "lastDataRow = LocalLimit") ||
		!strings.Contains(renamed, `Debug.Print value, lastDataRow, "lastRow", LocalLimit ' lastRow`) {
		t.Fatalf("local rename result:\n%s", renamed)
	}

	renamed = mustRenameSource(t, analyzer, doc, "Debug.Print value", "value", "newValue")
	if !strings.Contains(renamed, "Public Sub Sample(ByVal newValue As String)") ||
		!strings.Contains(renamed, "Debug.Print newValue, lastRow") {
		t.Fatalf("parameter rename result:\n%s", renamed)
	}

	renamed = mustRenameSource(t, analyzer, doc, "lastRow = LocalLimit", "LocalLimit", "MaxLocal")
	if !strings.Contains(renamed, "Const MaxLocal As Long = 3") ||
		!strings.Contains(renamed, "lastRow = MaxLocal") ||
		!strings.Contains(renamed, `Debug.Print value, lastRow, "lastRow", MaxLocal ' lastRow`) {
		t.Fatalf("const rename result:\n%s", renamed)
	}
}

func TestRenameDoesNotCrossProcedureScopes(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private Sub A()
    Dim x As Long
    x = 1
End Sub

Private Sub B()
    Dim x As Long
    x = 2
End Sub
`
	doc := renameDoc(t, source)

	renamed := mustRenameSource(t, analyzer, doc, "x = 1", "x", "valueA")
	if !strings.Contains(renamed, "Dim valueA As Long\n    valueA = 1") {
		t.Fatalf("procedure A local was not renamed:\n%s", renamed)
	}
	if !strings.Contains(renamed, "Dim x As Long\n    x = 2") {
		t.Fatalf("procedure B local should not be renamed:\n%s", renamed)
	}
}

func TestRenameParameterShadowsModuleVariable(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private value As Long

Private Sub A(ByVal value As Long)
    value = value + 1
End Sub

Private Sub B()
    value = 2
End Sub
`
	doc := renameDoc(t, source)

	renamed := mustRenameSource(t, analyzer, doc, "value = value + 1", "value", "argumentValue")
	if !strings.Contains(renamed, "Private value As Long") {
		t.Fatalf("module variable should not be renamed by parameter rename:\n%s", renamed)
	}
	if !strings.Contains(renamed, "Private Sub A(ByVal argumentValue As Long)\n    argumentValue = argumentValue + 1") {
		t.Fatalf("parameter should be renamed:\n%s", renamed)
	}
	if !strings.Contains(renamed, "Private Sub B()\n    value = 2") {
		t.Fatalf("module variable reference should remain:\n%s", renamed)
	}
}

func TestRenamePrivateModuleVariableConstAndProcedure(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private cache As Long
Private Const Limit As Long = 10

Private Sub Main()
    cache = Limit
    Call Helper
    Helper
End Sub

Private Sub Helper()
End Sub
`
	doc := renameDoc(t, source)

	renamed := mustRenameSource(t, analyzer, doc, "cache = Limit", "cache", "memo")
	if !strings.Contains(renamed, "Private memo As Long") || !strings.Contains(renamed, "memo = Limit") {
		t.Fatalf("module variable rename result:\n%s", renamed)
	}

	renamed = mustRenameSource(t, analyzer, doc, "cache = Limit", "Limit", "MaxRows")
	if !strings.Contains(renamed, "Private Const MaxRows As Long = 10") || !strings.Contains(renamed, "cache = MaxRows") {
		t.Fatalf("module const rename result:\n%s", renamed)
	}

	renamed = mustRenameSource(t, analyzer, doc, "Call Helper", "Helper", "DoHelp")
	if !strings.Contains(renamed, "Call DoHelp") ||
		!strings.Contains(renamed, "\n    DoHelp\n") ||
		!strings.Contains(renamed, "Private Sub DoHelp()") {
		t.Fatalf("private procedure rename result:\n%s", renamed)
	}
}

func TestRenameLabelReferences(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `Option Explicit
Private Sub Sample()
    On Error GoTo Cleanup
    GoSub Cleanup
    GoTo Cleanup
    Resume Cleanup
Cleanup:
    Debug.Print "Cleanup"
    Debug.Print "GoTo Cleanup"
End Sub
`
	doc := renameDoc(t, source)

	renamed := mustRenameSource(t, analyzer, doc, "GoTo Cleanup", "Cleanup", "Done")
	if !strings.Contains(renamed, "On Error GoTo Done") ||
		!strings.Contains(renamed, "GoSub Done") ||
		!strings.Contains(renamed, "GoTo Done") ||
		!strings.Contains(renamed, "Resume Done") ||
		!strings.Contains(renamed, "Done:") ||
		!strings.Contains(renamed, `Debug.Print "Cleanup"`) ||
		!strings.Contains(renamed, `Debug.Print "GoTo Cleanup"`) {
		t.Fatalf("label rename result:\n%s", renamed)
	}
}

func TestRenameRejectsReservedWords(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := renameDoc(t, `Option Explicit
Private Sub Sample()
    Dim value As Long
    value = 1
End Sub
`)

	for _, newName := range []string{"Dim", "Rem", "And", "Not"} {
		_, err := analyzer.Rename(doc, renamePosition(t, doc.Source, "value = 1", "value"), newName, []Document{doc}, nil)
		if err == nil || !strings.Contains(err.Error(), "invalid VBA identifier for rename") {
			t.Fatalf("Rename to reserved word %q error = %v", newName, err)
		}
	}
}

func TestRenameRejectsInScopeCollisions(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	doc := renameDoc(t, `Option Explicit
Private moduleValue As Long

Private Sub Sample(ByVal argumentValue As Long)
    Dim localValue As Long
    localValue = moduleValue + argumentValue
End Sub

Private Sub Other()
    Dim otherLocal As Long
    moduleValue = otherLocal
End Sub
`)

	cases := []struct {
		lineText string
		word     string
		newName  string
	}{
		{"localValue = moduleValue", "localValue", "argumentValue"},
		{"localValue = moduleValue", "localValue", "moduleValue"},
		{"moduleValue = otherLocal", "moduleValue", "otherLocal"},
	}
	for _, tc := range cases {
		_, err := analyzer.Rename(doc, renamePosition(t, doc.Source, tc.lineText, tc.word), tc.newName, []Document{doc}, nil)
		if err == nil || !strings.Contains(err.Error(), "in-scope symbol already exists") {
			t.Fatalf("Rename(%s -> %s) error = %v, want collision", tc.word, tc.newName, err)
		}
	}
}

func TestRenameRefusals(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	source := `VERSION 5.00
Begin VB.Form UserForm1
   Begin MSForms.TextBox txtName
   End
End
Attribute VB_Name = "UserForm1"
Option Explicit
Public Sub PublicApi()
    Range("A1").Value = 1
    missingValue = 1
End Sub
Private Sub CommandButton1_Click()
End Sub
`
	doc := Document{
		Path:       filepath.Join(t.TempDir(), "UserForm1.frm"),
		ModuleKind: "form",
		Source:     source,
	}

	assertRenameError(t, analyzer, doc, "Range(\"A1\").Value", "Value", "cannot rename external host member")
	assertRenameError(t, analyzer, doc, "missingValue = 1", "missingValue", "cannot rename unresolved identifier")
	assertRenameError(t, analyzer, doc, "Public Sub PublicApi", "PublicApi", "project-wide public rename is not supported yet")
	assertRenameError(t, analyzer, doc, "CommandButton1_Click", "CommandButton1_Click", "userform control/event rename is not supported yet")
	assertRenameError(t, analyzer, doc, "MSForms.TextBox txtName", "txtName", "userform control/event rename is not supported yet")
}

func TestRenameRejectsAmbiguousPublicSymbols(t *testing.T) {
	analyzer := newTestAnalyzer(t)
	root := t.TempDir()
	doc := Document{
		Path: filepath.Join(root, "src", "modules", "Main.bas"),
		Source: `Option Explicit
Private Sub Main()
    SharedName
End Sub
`,
	}
	open := []Document{
		doc,
		{Path: filepath.Join(root, "src", "modules", "A.bas"), Source: "Option Explicit\nPublic Sub SharedName()\nEnd Sub\n"},
		{Path: filepath.Join(root, "src", "modules", "B.bas"), Source: "Option Explicit\nPublic Sub SharedName()\nEnd Sub\n"},
	}

	_, err := analyzer.PrepareRename(doc, renamePosition(t, doc.Source, "SharedName", "SharedName"), open)
	if err == nil || !strings.Contains(err.Error(), "cannot rename ambiguous symbol") {
		t.Fatalf("ambiguous rename error = %v", err)
	}
}

func mustRenameSource(t *testing.T, analyzer Analyzer, doc Document, lineText, word, newName string) string {
	t.Helper()
	pos := renamePosition(t, doc.Source, lineText, word)
	edits, err := analyzer.Rename(doc, pos, newName, []Document{doc}, nil)
	if err != nil {
		t.Fatalf("Rename(%s -> %s) error = %v", word, newName, err)
	}
	return applyRenameEdits(t, doc.Source, edits)
}

func assertRenameError(t *testing.T, analyzer Analyzer, doc Document, lineText, word, want string) {
	t.Helper()
	_, err := analyzer.PrepareRename(doc, renamePosition(t, doc.Source, lineText, word), []Document{doc})
	if err == nil || !strings.Contains(err.Error(), want) {
		t.Fatalf("PrepareRename(%s) error = %v, want %q", word, err, want)
	}
}

func renameDoc(t *testing.T, source string) Document {
	t.Helper()
	return Document{
		Path:       filepath.Join(t.TempDir(), "Main.bas"),
		ModuleKind: "standard",
		Source:     source,
	}
}

func renamePosition(t *testing.T, source, lineText, word string) Position {
	t.Helper()
	for lineNo, line := range normalizedLines(source) {
		lineIdx := strings.Index(line, lineText)
		if lineIdx < 0 {
			continue
		}
		wordIdx := strings.Index(line[lineIdx:], word)
		if wordIdx < 0 {
			continue
		}
		start := lineIdx + wordIdx
		return Position{Line: lineNo, Character: utf16Len(line[:start+1])}
	}
	t.Fatalf("word %q on line containing %q not found", word, lineText)
	return Position{}
}

func applyRenameEdits(t *testing.T, source string, edits []RenameEdit) string {
	t.Helper()
	type byteEdit struct {
		start int
		end   int
		text  string
	}
	var converted []byteEdit
	for _, edit := range edits {
		converted = append(converted, byteEdit{
			start: byteOffsetForPosition(source, edit.Range.Start),
			end:   byteOffsetForPosition(source, edit.Range.End),
			text:  edit.NewText,
		})
	}
	sort.SliceStable(converted, func(i, j int) bool {
		return converted[i].start > converted[j].start
	})
	out := source
	for _, edit := range converted {
		out = out[:edit.start] + edit.text + out[edit.end:]
	}
	return out
}
