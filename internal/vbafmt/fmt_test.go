package vbafmt

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
)

func TestFormatBasIdempotent(t *testing.T) {
	input := `Option Explicit
Public Sub Main()
x = 1
If x = 1 Then
x = 2
End If
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if got == input {
		t.Fatal("expected formatting change for unformatted input")
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasPreserveTrailingBlankLine(t *testing.T) {
	input := "Option Explicit\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.HasSuffix(got, "\n") {
		t.Fatalf("expected trailing newline:\n%q", got)
	}
}

func TestFormatBasFixIndent(t *testing.T) {
	input := `Option Explicit

Sub Main()
x = 1
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "    x = 1") || !strings.Contains(got, "Sub Main") {
		t.Fatalf("expected fixed indent:\n%s", got)
	}
}

func TestFormatBasPreserveStringKeywords(t *testing.T) {
	input := `Sub Main()
    Dim s As String
    s = "If this is a string then End If not real"
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.Count(line, "If") > 2 {
			t.Fatalf("string keyword may have affected indent:\n%s", got)
		}
	}
	wantIndent := map[string]int{
		`s = "If this is a string then End If not real"`: 4,
		"End Sub": 0,
	}
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if want, ok := wantIndent[trimmed]; ok {
			gotIndent := len(line) - len(strings.TrimLeft(line, " "))
			if gotIndent != want {
				t.Fatalf("expected indent %d for %q, got %d:\n%s", want, trimmed, gotIndent, got)
			}
		}
	}
}

func TestFormatBasNestedSelectCase(t *testing.T) {
	input := `Sub Test()
Select Case x
Case 1
y = 1
Case 2
y = 2
Case Else
y = 0
End Select
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.TrimSpace(line) == "End Select" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected End Select at indent 4:\n%s\nline: %q indent: %d", got, line, indent)
			}
		}
		if strings.TrimSpace(line) == "Case 1" || strings.TrimSpace(line) == "Case 2" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected Case at indent 4:\n%s\nline: %q indent: %d", got, line, indent)
			}
		}
	}
}

func TestFormatBasLineContinuationPreserved(t *testing.T) {
	input := `Sub Main() _
    & "hello"
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "_") {
		t.Fatalf("line continuation missing:\n%s", got)
	}

	continuedIf := `Sub Main()
If condition Then _
    condition = False
value = 1
End If
End Sub
`
	continuedIfGot, err := FormatText(continuedIf, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(continuedIfGot, "Then _") {
		t.Fatalf("continued If opener missing continuation:\n%s", continuedIfGot)
	}
	if !strings.Contains(continuedIfGot, "End If") {
		t.Fatalf("continued If block missing End If:\n%s", continuedIfGot)
	}
	assertTrimmedLineIndent(t, continuedIfGot, "value = 1", 8)
	second, err := FormatText(continuedIfGot, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != continuedIfGot {
		t.Fatalf("continued If format not idempotent:\nfirst:\n%s\nsecond:\n%s", continuedIfGot, second)
	}
}

func assertTrimmedLineIndent(t *testing.T, text, trimmed string, want int) {
	t.Helper()
	for _, line := range strings.Split(strings.TrimRight(text, "\n"), "\n") {
		if strings.TrimSpace(line) != trimmed {
			continue
		}
		got := len(line) - len(strings.TrimLeft(line, " "))
		if got != want {
			t.Fatalf("expected %q indent %d, got %d:\n%s", trimmed, want, got, text)
		}
		return
	}
	t.Fatalf("line %q not found:\n%s", trimmed, text)
}

func TestFormatBasRemoveTrailingWhitespace(t *testing.T) {
	input := "Sub Main()   \t  \n    x = 1   \nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	for _, line := range strings.Split(strings.TrimRight(got, "\n"), "\n") {
		if strings.HasSuffix(line, " ") || strings.HasSuffix(line, "\t") {
			t.Fatalf("trailing whitespace not removed:\n%q", line)
		}
	}
}

func TestFormatClsPreservesAttributes(t *testing.T) {
	input := `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Sub DoSomething()
    x = 1
End Sub
`
	got, err := FormatText(input, true)
	if err != nil {
		t.Fatal(err)
	}
	for _, attr := range []string{
		"VERSION 1.0 CLASS",
		"BEGIN",
		"MultiUse = -1  'True",
		"END",
		"Attribute VB_Name",
		"Attribute VB_GlobalNameSpace",
		"Attribute VB_Creatable",
		"Attribute VB_PredeclaredId",
		"Attribute VB_Exposed",
	} {
		if !strings.Contains(got, attr) {
			t.Fatalf("missing cls header line %q:\n%s", attr, got)
		}
	}
	if strings.Contains(got, "    BEGIN") || strings.Contains(got, "    END") {
		t.Fatalf("class metadata block BEGIN/END should not be indented:\n%s", got)
	}
	if strings.Contains(got, "    MultiUse") {
		t.Fatalf("class metadata MultiUse should not be re-indented:\n%s", got)
	}
}

func TestFormatClsPreservesAdditionalBeginEndProperties(t *testing.T) {
	input := "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\n  Persistable = 0  'NotPersistable\r\nEND\r\nAttribute VB_Name = \"ThisWorkbook\"\r\nOption Explicit\r\nPublic Sub Hello()\r\nx = 1\r\nEnd Sub\r\n"
	got, err := FormatText(input, true)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"VERSION 1.0 CLASS",
		"BEGIN",
		"MultiUse = -1  'True",
		"Persistable = 0  'NotPersistable",
		"END",
		"Attribute VB_Name = \"ThisWorkbook\"",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("missing cls header line %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "    Persistable") {
		t.Fatalf("Persistable inside BEGIN/END should not be re-indented:\n%s", got)
	}
	second, err := FormatText(got, true)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatRemComment(t *testing.T) {
	input := `Sub Main()
Rem this is a comment
x = 1
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Rem this is a comment") {
		t.Fatalf("Rem comment not preserved:\n%s", got)
	}
}

func TestFormatBasForEachLoop(t *testing.T) {
	input := `Sub Main()
Dim v As Variant
For Each v In Array(1, 2, 3)
x = v
Next v
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.TrimSpace(line) == "Next v" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected Next at indent 4:\n%s", got)
			}
		}
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for For Each:\n%s\n%s", got, second)
	}
}

func TestFormatBasWithBlock(t *testing.T) {
	input := `Sub Main()
With Sheet1
.Cells(1, 1) = "hello"
End With
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for With block:\n%s\n%s", got, second)
	}
}

func TestFormatBasTypeEnum(t *testing.T) {
	input := `Type Point
x As Long
y As Long
End Type

Enum Color
Red
Green
Blue
End Enum
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Type/Enum:\n%s\n%s", got, second)
	}
}

func TestFormatBasDoWhileLoop(t *testing.T) {
	input := `Sub Main()
Do While x < 10
x = x + 1
Loop
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Do While:\n%s\n%s", got, second)
	}
}

func TestFormatBasIfElseIdempotent(t *testing.T) {
	input := `Sub Main()
If x = 1 Then
y = 1
ElseIf x = 2 Then
y = 2
Else
y = 0
End If
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for If/Else/ElseIf:\n%s\n%s", got, second)
	}
}

func TestFormatBasElseBodyIndent(t *testing.T) {
	input := `Sub Main()
If x = 1 Then
y = 1
Else
y = 0
End If
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		switch trimmed {
		case "Sub Main()", "End Sub":
			// top-level
		case "If x = 1 Then", "Else", "End If":
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected %q at indent 4, got indent %d:\n%s", trimmed, indent, got)
			}
		case "y = 1", "y = 0":
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 8 {
				t.Fatalf("expected %q at indent 8 (body under If/Else), got indent %d:\n%s", trimmed, indent, got)
			}
		}
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for If/Else body indent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasWend(t *testing.T) {
	input := `Sub Main()
While x < 10
x = x + 1
Wend
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for While/Wend:\n%s\n%s", got, second)
	}
}

func TestFormatBasConditionalCompilation(t *testing.T) {
	input := `Sub Main()
#If Win64 Then
x = 1
#Else
x = 0
#End If
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for #If/#Else/#End If:\n%s\n%s", got, second)
	}
	assertTrimmedLineIndent(t, got, "#If Win64 Then", 4)
	assertTrimmedLineIndent(t, got, "x = 1", 4)
	assertTrimmedLineIndent(t, got, "#Else", 4)
	assertTrimmedLineIndent(t, got, "x = 0", 4)
	assertTrimmedLineIndent(t, got, "#End If", 4)
}

func TestFormatBasProperty(t *testing.T) {
	input := `Property Get Value() As Long
Value = m_Value
End Property

Property Let Value(v As Long)
m_Value = v
End Property

Property Set Ref(obj As Object)
Set m_Ref = obj
End Property
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Property:\n%s\n%s", got, second)
	}
}

func TestFormatBasFunction(t *testing.T) {
	input := `Function Add(a As Long, b As Long) As Long
Add = a + b
End Function
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Function:\n%s\n%s", got, second)
	}
}

func TestFormatBasBlankLineNormalization(t *testing.T) {
	input := "Sub Main()\n\n\n\n\nx = 1\n\n\n\n\nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	consecutiveBlanks := 0
	maxBlanks := 0
	for _, line := range lines {
		if strings.TrimSpace(line) == "" {
			consecutiveBlanks++
			if consecutiveBlanks > maxBlanks {
				maxBlanks = consecutiveBlanks
			}
		} else {
			consecutiveBlanks = 0
		}
	}
	if maxBlanks > 1 {
		t.Fatalf("max 1 consecutive blank line expected, got %d:\n%s", maxBlanks, got)
	}
}

func TestFormatBasOptionExplicitGap(t *testing.T) {
	input := "Option Explicit\nSub Main()\nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Option Explicit\n\nSub Main") {
		t.Fatalf("expected blank line after Option Explicit:\n%s", got)
	}
}

func TestFormatBasProcedureGap(t *testing.T) {
	input := "Sub A()\nEnd Sub\nSub B()\nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "End Sub\n\nSub B") {
		t.Fatalf("expected blank line between procedures:\n%s", got)
	}
}

func TestFileResultDetectsChange(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	fr, err := formatFile(path, "")
	if err != nil {
		t.Fatal(err)
	}
	if !fr.Changed {
		t.Fatal("expected change detected")
	}
	if fr.Path != path {
		t.Fatalf("path = %q", fr.Path)
	}
}

func TestFileResultNoChangeWhenFormatted(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Formatted.bas")
	formatted := "Sub Main()\n    x = 1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(formatted), 0644); err != nil {
		t.Fatal(err)
	}
	fr, err := formatFile(path, "")
	if err != nil {
		t.Fatal(err)
	}
	if fr.Changed {
		t.Fatalf("expected no change for already formatted file:\noriginal:\n%q\nformatted:\n%q", fr.Original, fr.Formatted)
	}
}

func TestResolveExplicitPathsFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.bas")
	if err := os.WriteFile(path, []byte("Sub Main()\nEnd Sub\n"), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Root: dir, Paths: []string{path}}
	files, err := resolveExplicitPaths(opts)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 1 || files[0] != path {
		t.Fatalf("expected [%s], got %v", path, files)
	}
}

func TestResolveExplicitPathsDir(t *testing.T) {
	dir := t.TempDir()
	sub := filepath.Join(dir, "src", "modules")
	if err := os.MkdirAll(sub, 0755); err != nil {
		t.Fatal(err)
	}
	p1 := filepath.Join(sub, "A.bas")
	p2 := filepath.Join(sub, "B.cls")
	if err := os.WriteFile(p1, []byte("x"), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(p2, []byte("x"), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Root: dir, Paths: []string{filepath.Join(dir, "src")}}
	files, err := resolveExplicitPaths(opts)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 2 {
		t.Fatalf("expected 2 files, got %v", files)
	}
}

func TestDiffProducesUnifiedOutput(t *testing.T) {
	orig := "Sub Main()\nx = 1\nEnd Sub\n"
	fmtStr := "Sub Main()\n    x = 1\nEnd Sub\n"
	diff := Diff("test.bas", orig, fmtStr)
	if !strings.Contains(diff, "@@") || !strings.Contains(diff, "x = 1") {
		t.Fatalf("expected unified diff:\n%s", diff)
	}
}

func TestDiffNoChange(t *testing.T) {
	same := "Sub Main()\nEnd Sub\n"
	diff := Diff("test.bas", same, same)
	if diff != "" {
		t.Fatalf("expected empty diff for unchanged text")
	}
}

func TestFormatBasDoUntil(t *testing.T) {
	input := `Sub Main()
Do Until x > 10
x = x + 1
Loop
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Do Until:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.TrimSpace(line) == "Loop" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected Loop at indent 4:\n%s", got)
			}
		}
	}
}

func TestFormatBasLoopWhile(t *testing.T) {
	input := `Sub Main()
Do
x = x + 1
Loop While x < 10
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Loop While:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasLoopUntil(t *testing.T) {
	input := `Sub Main()
Do
x = x + 1
Loop Until x >= 10
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Loop Until:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasConditionalCompilationElseIf(t *testing.T) {
	input := `Sub Main()
#If Win64 Then
x = 1
#ElseIf Win32 Then
x = 0
#Else
x = -1
#End If
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for #If/#ElseIf/#Else/#End If:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
	assertTrimmedLineIndent(t, got, "#If Win64 Then", 4)
	assertTrimmedLineIndent(t, got, "x = 1", 4)
	assertTrimmedLineIndent(t, got, "#ElseIf Win32 Then", 4)
	assertTrimmedLineIndent(t, got, "x = 0", 4)
	assertTrimmedLineIndent(t, got, "#Else", 4)
	assertTrimmedLineIndent(t, got, "x = -1", 4)
	assertTrimmedLineIndent(t, got, "#End If", 4)
}

func TestResolveProjectFilesIncludesSidecarFormCode(t *testing.T) {
	dir := t.TempDir()
	if err := os.MkdirAll(filepath.Join(dir, "src", "forms", "code"), 0755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "src", "forms", "code", "UserForm1.bas")
	if err := os.WriteFile(path, []byte("Option Explicit\n"), 0644); err != nil {
		t.Fatal(err)
	}
	files, err := resolveProjectFiles(FmtOptions{Root: dir})
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range files {
		if filepath.Clean(file) == filepath.Clean(path) {
			return
		}
	}
	t.Fatalf("expected sidecar form code file in default project files: %v", files)
}

func TestFormatBasPrivateStaticSub(t *testing.T) {
	input := `Private Static Sub Main()
x = 1
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for Private Static Sub:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.TrimSpace(line) == "x = 1" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected body at indent 4 for Private Static Sub:\n%s", got)
			}
		}
	}
}

func TestFormatRemCommentNoSpace(t *testing.T) {
	input := `Sub Main()
Remcomment
x = 1
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Remcomment") {
		t.Fatalf("Rem without space not preserved:\n%s", got)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		if strings.TrimSpace(line) == "Remcomment" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("expected Remcomment (not a VBA comment) at indent 4:\n%s", got)
			}
		}
	}
}

func TestFormatFileSkipsUnsupportedExtension(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "UserForm1.frm")
	if err := os.WriteFile(path, []byte("VERSION 5.00\nBegin Form\nEnd\n"), 0644); err != nil {
		t.Fatal(err)
	}
	fr, err := formatFile(path, "")
	if err != nil {
		t.Fatal(err)
	}
	if !fr.Skipped {
		t.Fatal("expected .frm file to be skipped")
	}
	if fr.SkipReason == "" {
		t.Fatal("expected non-empty skip reason")
	}
}

func TestResolveExplicitPathsIncludesFrm(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "UserForm1.frm")
	if err := os.WriteFile(path, []byte("VERSION 5.00\n"), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Root: dir, Paths: []string{path}}
	files, err := resolveExplicitPaths(opts)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 1 {
		t.Fatalf("expected 1 file for explicit .frm path, got %v", files)
	}
}

func TestResolveExplicitPathsCrossArgDedup(t *testing.T) {
	dir := t.TempDir()
	sub1 := filepath.Join(dir, "sub1")
	sub2 := filepath.Join(dir, "sub2")
	if err := os.MkdirAll(sub1, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(sub2, 0755); err != nil {
		t.Fatal(err)
	}
	shared := filepath.Join(sub1, "shared.bas")
	if err := os.WriteFile(shared, []byte("Sub Foo()\nEnd Sub\n"), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(sub1, "a.bas"), []byte("Sub A()\nEnd Sub\n"), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(sub2, "b.bas"), []byte("Sub B()\nEnd Sub\n"), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Root: dir, Paths: []string{sub1, sub2}}
	files, err := resolveExplicitPaths(opts)
	if err != nil {
		t.Fatal(err)
	}
	if len(files) != 3 {
		t.Fatalf("expected 3 unique files, got %d: %v", len(files), files)
	}
	optsDedup := FmtOptions{Root: dir, Paths: []string{sub1, sub1, sub2}}
	filesDedup, err := resolveExplicitPaths(optsDedup)
	if err != nil {
		t.Fatal(err)
	}
	if len(filesDedup) != 3 {
		t.Fatalf("expected 3 unique files with duplicate dir args, got %d: %v", len(filesDedup), filesDedup)
	}
}

func TestFormatTextEmptyInput(t *testing.T) {
	got, err := FormatText("", false)
	if err != nil {
		t.Fatal(err)
	}
	if got != "" {
		t.Fatalf("expected empty output for empty input, got %q", got)
	}
}

func TestFormatTextStdinCls(t *testing.T) {
	input := `VERSION 1.0 CLASS
BEGIN
MultiUse = -1
END
Attribute VB_Name = "Test"
Option Explicit

Public Sub DoSomething()
x = 1
End Sub
`
	got, err := FormatText(input, true)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "VERSION 1.0 CLASS") {
		t.Fatalf("cls formatting failed:\n%s", got)
	}
	second, err := FormatText(got, true)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("cls format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestSummarizeResultsSkippedNotCountedAsUnchanged(t *testing.T) {
	results := []FileResult{
		{Path: "a.bas", Changed: true, Formatted: "formatted a"},
		{Path: "b.bas", Changed: false},
		{Path: "c.frm", Skipped: true, SkipReason: "unsupported extension: .frm"},
	}
	opts := FmtOptions{}
	r, err := summarizeResults(results, opts)
	if err != nil {
		t.Fatal(err)
	}
	if r.Total != 3 {
		t.Fatalf("expected Total=3, got %d", r.Total)
	}
	if r.Changed != 1 {
		t.Fatalf("expected Changed=1, got %d", r.Changed)
	}
	if r.Unchanged != 1 {
		t.Fatalf("expected Unchanged=1 (skipped files should not count as unchanged), got %d", r.Unchanged)
	}
	if r.Skipped != 1 {
		t.Fatalf("expected Skipped=1, got %d", r.Skipped)
	}
	if len(r.FormattedByPath) != 1 {
		t.Fatalf("expected 1 entry in FormattedByPath, got %d", len(r.FormattedByPath))
	}
	if got := r.FormattedByPath["a.bas"]; got != "formatted a" {
		t.Fatalf("expected FormattedByPath[a.bas]=%q, got %q", "formatted a", got)
	}
}

func TestRunWriteWritesFiles(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Write: true, Paths: []string{path}, Root: dir}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Changed != 1 {
		t.Fatalf("expected 1 changed, got %d", result.Changed)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) == original {
		t.Fatal("expected file to be written with formatted content")
	}
}

func TestRunCheckOnFormattedFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.bas")
	formatted := "Sub Main()\n    x = 1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(formatted), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Check: true, Paths: []string{path}, Root: dir}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Changed != 0 {
		t.Fatalf("expected 0 changed for already formatted file, got %d", result.Changed)
	}
}

func TestRunCheckDetectsUnformatted(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Check: true, Paths: []string{path}, Root: dir}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Changed != 1 {
		t.Fatalf("expected 1 changed for unformatted file, got %d", result.Changed)
	}
}

func TestRunDiffNoWrite(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "test.bas")
	original := "Sub Main()\nx=1\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(original), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Diff: true, Paths: []string{path}, Root: dir}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Changed != 1 {
		t.Fatalf("expected 1 changed for unformatted file, got %d", result.Changed)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(data) != original {
		t.Fatal("diff mode should not write files")
	}
}

func TestRunExplicitFrmProducesSkippedResult(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "UserForm1.frm")
	if err := os.WriteFile(path, []byte("VERSION 5.00\n"), 0644); err != nil {
		t.Fatal(err)
	}
	opts := FmtOptions{Paths: []string{path}, Root: dir}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Skipped != 1 {
		t.Fatalf("expected 1 skipped for .frm path, got %d", result.Skipped)
	}
	if result.Total != 1 {
		t.Fatalf("expected Total=1 for explicit .frm path, got %d", result.Total)
	}
	if len(result.SkippedPaths) != 1 {
		t.Fatal("expected SkippedPaths to contain .frm path")
	}
	if len(result.SkippedReasons) != 1 {
		t.Fatal("expected SkippedReasons to contain skip reason")
	}
}

func TestRunEmptyProjectProducesZeroTotal(t *testing.T) {
	dir := t.TempDir()
	opts := FmtOptions{Root: dir, Cfg: config.Config{
		Src: config.SourceConfig{
			Modules:  "nonexistent/modules",
			Classes:  "nonexistent/classes",
			Workbook: "nonexistent/workbook",
		},
	}}
	result, err := Run(opts)
	if err != nil {
		t.Fatal(err)
	}
	if result.Total != 0 {
		t.Fatalf("expected Total=0 for empty project, got %d", result.Total)
	}
}

func TestFormatClsPreservesBlankLinesInHeader(t *testing.T) {
	input := `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False

Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Sub DoSomething()
    x = 1
End Sub
`
	got, err := FormatText(input, true)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(got, "\n")
	inHeader := true
	for _, line := range lines {
		if strings.TrimSpace(line) == "Option Explicit" {
			inHeader = false
		}
		if inHeader && strings.TrimSpace(line) == "" {
			continue
		}
		if inHeader {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 0 && !strings.HasPrefix(strings.TrimSpace(line), "MultiUse") {
				t.Fatalf("cls header line should not be re-indented: %q\n%s", line, got)
			}
		}
	}
	second, err := FormatText(got, true)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("cls format not idempotent with blank lines in header:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasOneLineIfNotExpanded(t *testing.T) {
	input := `Sub Main()
If x Then y = 1
z = 2
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "If x Then y = 1" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("single-line If should be at indent 4, got %d:\n%s", indent, got)
			}
		}
		if trimmed == "z = 2" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("line after single-line If should be at indent 4, got %d:\n%s", indent, got)
			}
		}
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("single-line If format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasColonSeparatedNotSplit(t *testing.T) {
	input := `Sub Main()
If x Then y = 1: z = 2
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "If x Then y = 1: z = 2") {
		t.Fatalf("colon-separated statement should not be split:\n%s", got)
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("colon-separated format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasSelectCaseBodyIndentConsistency(t *testing.T) {
	input := `Sub Test()
Select Case x
Case 1
y = 1
z = 2
Case 2
y = 3
z = 4
Case Else
y = 0
z = -1
End Select
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	bodyLines := map[string]int{
		"y = 1":  8,
		"z = 2":  8,
		"y = 3":  8,
		"z = 4":  8,
		"y = 0":  8,
		"z = -1": 8,
	}
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if want, ok := bodyLines[trimmed]; ok {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != want {
				t.Fatalf("expected %q at indent %d, got %d:\n%s", trimmed, want, indent, got)
			}
		}
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent for multi-case Select:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatBasColonClosedBlockNotIndented(t *testing.T) {
	input := `Sub Main()
Do While i < 10: i = i + 1: Loop
j = 1
End Sub
`
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "Do While i < 10: i = i + 1: Loop" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("colon-closed Do block should be at indent 4, got %d:\n%s", indent, got)
			}
		}
		if trimmed == "j = 1" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("line after colon-closed block should be at indent 4, got %d:\n%s", indent, got)
			}
		}
		if trimmed == "End Sub" {
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 0 {
				t.Fatalf("End Sub should be at indent 0, got %d:\n%s", indent, got)
			}
		}
	}
	second, err := FormatText(got, false)
	if err != nil {
		t.Fatal(err)
	}
	if second != got {
		t.Fatalf("format not idempotent:\nfirst:\n%s\nsecond:\n%s", got, second)
	}
}

func TestFormatTextPreservesExistingLineNumbers(t *testing.T) {
	input := "Sub Main()\n10 x=1\n20 If x = 1 Then\n30 y = 2\n40 End If\nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"10      x=1",
		"20      If x = 1 Then",
		"30          y = 2",
		"40      End If",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected preserved line number %q:\n%s", want, got)
		}
	}
}

func TestFormatTextWithLineNumbersAdd(t *testing.T) {
	input := "Public Sub Sample()\n    Dim x As Integer\n    x = 1 / 0\nEnd Sub\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"10      Dim x As Integer",
		"20      x = 1 / 0",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected numbered line %q:\n%s", want, got)
		}
	}
	if strings.Contains(got, "10  Public Sub Sample()") || strings.Contains(got, "30  End Sub") {
		t.Fatalf("procedure boundary should not be numbered:\n%s", got)
	}
}

func TestFormatTextWithLineNumbersAddSkipsSelectCaseStructuralLines(t *testing.T) {
	input := "Private Function IsUnreservedUrlByte(ByVal byteValue As Long) As Boolean\n    Select Case byteValue\n    Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126\n        IsUnreservedUrlByte = True\n    Case Else\n        IsUnreservedUrlByte = False\n    End Select\nEnd Function\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"    Select Case byteValue",
		"    Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126",
		"    Case Else",
		"    End Select",
		"10          IsUnreservedUrlByte = True",
		"20          IsUnreservedUrlByte = False",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected output %q:\n%s", want, got)
		}
	}
	for _, unwanted := range []string{
		"10      Select Case byteValue",
		"10      Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126",
		"20      Case Else",
		"30      End Select",
	} {
		if strings.Contains(got, unwanted) {
			t.Fatalf("select/case structural line should not be numbered %q:\n%s", unwanted, got)
		}
	}
}

func TestFormatTextWithLineNumbersAddIsIdempotent(t *testing.T) {
	input := "Public Sub Sample()\n    Dim x As Integer\n    x = 1 / 0\nEnd Sub\n"
	first, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatTextWithOptions(first, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if second != first {
		t.Fatalf("line-number add should be idempotent:\nfirst:\n%s\nsecond:\n%s", first, second)
	}
}

func TestFormatTextWithLineNumbersAddSelectCaseIsIdempotent(t *testing.T) {
	input := "Private Function IsUnreservedUrlByte(ByVal byteValue As Long) As Boolean\n    Select Case byteValue\n    Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126\n        IsUnreservedUrlByte = True\n    End Select\nEnd Function\n"
	first, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatTextWithOptions(first, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if second != first {
		t.Fatalf("select case line-number add should be idempotent:\nfirst:\n%s\nsecond:\n%s", first, second)
	}
}

func TestFormatTextWithLineNumbersAddNumbersOnlyFirstLineOfContinuation(t *testing.T) {
	input := "Public Sub LogMessage(ByVal Message As String)\n    payload = \"{\" & _\n        JsonProperty(\"event\", \"debug_log\") & \",\" & _\n        JsonProperty(\"message\", Message) & \",\" & _\n        JsonProperty(\"runtime_mode\", XlflowRuntime.ModeName()) & \",\" & _\n        JsonProperty(\"source\", \"XlflowDebug.Log\") & \"}\"\nEnd Sub\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"10      payload = \"{\" & _",
		"    JsonProperty(\"event\", \"debug_log\") & \",\" & _",
		"    JsonProperty(\"message\", Message) & \",\" & _",
		"    JsonProperty(\"runtime_mode\", XlflowRuntime.ModeName()) & \",\" & _",
		"    JsonProperty(\"source\", \"XlflowDebug.Log\") & \"}\"",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected continued output %q:\n%s", want, got)
		}
	}
	for _, unwanted := range []string{
		"20      JsonProperty(\"event\", \"debug_log\") & \",\" & _",
		"30      JsonProperty(\"message\", Message) & \",\" & _",
		"40      JsonProperty(\"runtime_mode\", XlflowRuntime.ModeName()) & \",\" & _",
		"50      JsonProperty(\"source\", \"XlflowDebug.Log\") & \"}\"",
	} {
		if strings.Contains(got, unwanted) {
			t.Fatalf("continuation tail should not be numbered %q:\n%s", unwanted, got)
		}
	}
}

func TestFormatTextWithLineNumbersAddContinuationIsIdempotent(t *testing.T) {
	input := "Public Sub LogMessage(ByVal Message As String)\n    payload = \"{\" & _\n        JsonProperty(\"event\", \"debug_log\") & \",\" & _\n        JsonProperty(\"message\", Message) & \",\" & _\n        JsonProperty(\"runtime_mode\", XlflowRuntime.ModeName()) & \",\" & _\n        JsonProperty(\"source\", \"XlflowDebug.Log\") & \"}\"\nEnd Sub\n"
	first, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	second, err := FormatTextWithOptions(first, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if second != first {
		t.Fatalf("continuation line-number add should be idempotent:\nfirst:\n%s\nsecond:\n%s", first, second)
	}
}

func TestFormatTextWithLineNumbersRemove(t *testing.T) {
	input := "Public Sub Sample()\n10      Dim x As Integer\n20      x = 1 / 0\nEnd Sub\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeRemove})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(got, "10      ") || strings.Contains(got, "20      ") {
		t.Fatalf("line numbers should be removed:\n%s", got)
	}
	for _, want := range []string{"Dim x As Integer", "x = 1 / 0"} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected preserved body line %q:\n%s", want, got)
		}
	}
}

func TestFormatTextWithLineNumbersRenumber(t *testing.T) {
	input := "Public Sub Sample()\n100     Dim x As Integer\n700     x = 1 / 0\nEnd Sub\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeRenumber})
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		"10      Dim x As Integer",
		"20      x = 1 / 0",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected renumbered line %q:\n%s", want, got)
		}
	}
}

func TestFormatTextWithLineNumbersAddPreservesCommentsAndAttributes(t *testing.T) {
	input := "Attribute VB_Name = \"Main\"\nOption Explicit\n\nPublic Sub Sample()\n    ' comment\n    Dim x As Integer\n\n    x = 1 / 0\nEnd Sub\n"
	got, err := FormatTextWithOptions(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(got, "Attribute VB_Name = \"Main\"") {
		t.Fatalf("expected Attribute line preserved:\n%s", got)
	}
	if !strings.Contains(got, "    ' comment") {
		t.Fatalf("expected comment preserved without numbering:\n%s", got)
	}
	if strings.Contains(got, "10      ' comment") {
		t.Fatalf("comment line should not be numbered:\n%s", got)
	}
}

func TestFormatTextWithLineNumbersSkipsAmbiguousNumericLabels(t *testing.T) {
	input := "Public Sub Sample()\n10      Dim x As Integer\n20      GoTo 10\nEnd Sub\n"
	got, detail, err := formatTextDetailed(input, false, FormatConfig{LineNumbers: LineNumberModeRemove})
	if err != nil {
		t.Fatal(err)
	}
	if len(detail.Warnings) == 0 {
		t.Fatal("expected warning for ambiguous numeric labels")
	}
	if got != FormatMust(t, input) {
		t.Fatalf("ambiguous numeric labels should be left unchanged except normal formatting:\n%s", got)
	}
}

func TestFormatTextWithLineNumbersIgnoresGoToInsideStringLiteral(t *testing.T) {
	input := "Public Sub Sample()\n    MsgBox \"GoTo 10\"\n    x = 1\nEnd Sub\n"
	got, detail, err := formatTextDetailed(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if len(detail.Warnings) != 0 {
		t.Fatalf("did not expect numeric-label warning for string literal: %#v", detail.Warnings)
	}
	for _, want := range []string{
		"10      MsgBox \"GoTo 10\"",
		"20      x = 1",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected numbered output %q:\n%s", want, got)
		}
	}
}

func TestFormatTextWithLineNumbersAddSkipsLegacyNumberedComment(t *testing.T) {
	input := "10  ' legacy comment\nPublic Sub Sample()\n    x = 1\nEnd Sub\n"
	got, detail, err := formatTextDetailed(input, false, FormatConfig{LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if len(detail.Warnings) == 0 {
		t.Fatal("expected warning for legacy numbered non-executable line")
	}
	if got != FormatMust(t, input) {
		t.Fatalf("legacy numbered comment should prevent add rewrite:\n%s", got)
	}
}

func TestFormatTextWithLineNumbersRenumberIncludesLegacyNumberedComment(t *testing.T) {
	input := "10  ' legacy comment\n20  Public Sub Sample()\n30      x = 1\nEnd Sub\n"
	got, detail, err := formatTextDetailed(input, false, FormatConfig{LineNumbers: LineNumberModeRenumber})
	if err != nil {
		t.Fatal(err)
	}
	if len(detail.Warnings) != 0 {
		t.Fatalf("did not expect warning: %#v", detail.Warnings)
	}
	for _, want := range []string{
		"10  ' legacy comment",
		"20  Public Sub Sample()",
		"30      x = 1",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected renumbered output %q:\n%s", want, got)
		}
	}
	if strings.Count(got, "\n10  ") > 1 {
		t.Fatalf("renumber should not produce duplicate labels:\n%s", got)
	}
}

func TestFormatTextWithLineNumbersRenumberLegacyOnlyNumberedLines(t *testing.T) {
	input := "30  ' legacy comment a\n90  ' legacy comment b\n"
	got, detail, err := formatTextDetailed(input, false, FormatConfig{LineNumbers: LineNumberModeRenumber})
	if err != nil {
		t.Fatal(err)
	}
	if len(detail.Warnings) != 0 {
		t.Fatalf("did not expect warning: %#v", detail.Warnings)
	}
	for _, want := range []string{
		"10  ' legacy comment a",
		"20  ' legacy comment b",
	} {
		if !strings.Contains(got, want) {
			t.Fatalf("expected renumbered legacy-only line %q:\n%s", want, got)
		}
	}
	if !detail.Changed {
		t.Fatalf("expected renumber legacy-only lines to report changed")
	}
	if detail.LinesRenumbered != 2 {
		t.Fatalf("expected 2 renumbered lines, got %d", detail.LinesRenumbered)
	}
}

func TestRunLineNumberSummary(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "Sample.bas")
	input := "Public Sub Sample()\n    Dim x As Integer\n    x = 1 / 0\nEnd Sub\n"
	if err := os.WriteFile(path, []byte(input), 0644); err != nil {
		t.Fatal(err)
	}
	result, err := Run(FmtOptions{Paths: []string{path}, Root: dir, LineNumbers: LineNumberModeAdd})
	if err != nil {
		t.Fatal(err)
	}
	if result.LineNumbers.Mode != LineNumberModeAdd {
		t.Fatalf("mode = %q, want %q", result.LineNumbers.Mode, LineNumberModeAdd)
	}
	if result.LineNumbers.FilesChanged != 1 {
		t.Fatalf("files_changed = %d, want 1", result.LineNumbers.FilesChanged)
	}
	if result.LineNumbers.LinesAdded != 2 {
		t.Fatalf("lines_added = %d, want 2", result.LineNumbers.LinesAdded)
	}
}

func TestFormatBasLabelLineIsNotIndented(t *testing.T) {
	input := "Public Sub Sample()\n    On Error GoTo ErrorHandler\n    x = 1\n    Exit Sub\n    ErrorHandler:\n    x = 2\nEnd Sub\n"
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	lines := strings.Split(strings.TrimRight(got, "\n"), "\n")
	for _, line := range lines {
		switch strings.TrimSpace(line) {
		case "ErrorHandler:":
			if line != "ErrorHandler:" {
				t.Fatalf("label should not be indented:\n%s", got)
			}
		case "x = 2":
			indent := len(line) - len(strings.TrimLeft(line, " "))
			if indent != 4 {
				t.Fatalf("statement after label should remain in procedure indent:\n%s", got)
			}
		}
	}
}

func FormatMust(t *testing.T, input string) string {
	t.Helper()
	got, err := FormatText(input, false)
	if err != nil {
		t.Fatal(err)
	}
	return got
}
