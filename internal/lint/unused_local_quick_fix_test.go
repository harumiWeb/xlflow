package lint

import (
	"strings"
	"testing"

	"github.com/harumiWeb/xlflow/internal/config"
	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
)

func TestUnusedLocalQuickFix(t *testing.T) {
	for _, typ := range strings.Fields("Boolean Byte Integer Long LongLong LongPtr Single Double Currency Date String Variant") {
		t.Run(typ, func(t *testing.T) {
			source := "Attribute VB_Name = \"Main\"\r\nOption Explicit\r\nSub Work()\r\n    Dim scratch As " + typ + " ' 保持😀\r\n    Debug.Print \"scratch\" ' scratch\r\nEnd Sub\r\n"
			linter, doc := scratchLocalQuickFixDocument(t, source)
			fixes, err := linter.QuickFixesParsedContext(t.Context(), doc)
			if err != nil || len(fixes) != 1 {
				_ = doc.ReadContext(t.Context(), func(view vbaast.ParsedView) error { t.Log(view.Root.ToSexp()); return nil })
				t.Fatalf("fixes = %+v, error = %v", fixes, err)
			}
			fix := fixes[0]
			if fix.Title != "Remove unused local variable" || fix.Diagnostic.Code != "VB020" || len(fix.Edits) != 1 {
				t.Fatalf("fix = %+v", fix)
			}
			edit := fix.Edits[0]
			got := source[:edit.StartByte] + edit.NewText + source[edit.EndByte:]
			want := strings.Replace(source, "Dim scratch As "+typ, "", 1)
			if got != want {
				t.Fatalf("edited source = %q, want %q", got, want)
			}
			issues, err := linter.LintParsed(doc)
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, issue := range issues {
				if issue.Code == "VB020" {
					found = true
					if issue != fix.Diagnostic {
						t.Fatalf("diagnostic = %+v, want %+v", fix.Diagnostic, issue)
					}
				}
			}
			if !found {
				t.Fatal("missing VB020 diagnostic")
			}
		})
	}
}

func TestUnusedLocalQuickFixWithholdsUnsafeDeclarations(t *testing.T) {
	for name, body := range map[string]string{
		"read":                  "Dim scratch As Long\nDebug.Print scratch",
		"write":                 "Dim scratch As Long\nscratch = 1",
		"case write":            "Dim scratch As Long\nSCRATCH = 1",
		"bracket write":         "Dim scratch As Long\n[scratch] = 1",
		"suffix write":          "Dim scratch As Long\nscratch& = 1",
		"input":                 "Dim scratch As Long\nInput #1, scratch",
		"loop":                  "Dim scratch As Long\nFor scratch = 1 To 3\nNext",
		"redim":                 "Dim scratch As Variant\nReDim scratch(2)",
		"duplicate":             "Dim scratch As Long\nDim scratch As Long",
		"const":                 "Const scratch As Long = 1",
		"static":                "Static scratch As Long",
		"new":                   "Dim scratch As New Collection",
		"object":                "Dim scratch As Object",
		"class":                 "Dim scratch As Collection",
		"udt":                   "Dim scratch As UserType",
		"array":                 "Dim scratch() As Long",
		"fixed string":          "Dim scratch As String * 5",
		"implicit":              "Dim scratch",
		"suffix":                "Dim scratch&",
		"bracket declaration":   "Dim [scratch] As Long",
		"multiple":              "Dim scratch As Long, other As Long",
		"colon after":           "Dim scratch As Long: Debug.Print 1",
		"colon before":          "Debug.Print 1: Dim scratch As Long",
		"label":                 "here: Dim scratch As Long",
		"continued":             "Dim scratch _\n As Long",
		"conditional":           "#If VBA7 Then\nDim scratch As Long\n#End If",
		"conditional elsewhere": "Dim scratch As Long\n#If VBA7 Then\nDebug.Print 1\n#End If",
		"malformed":             "Dim scratch As Long\nDebug.Print (",
		"ignored":               "Dim unused As Long",
	} {
		t.Run(name, func(t *testing.T) {
			linter, doc := scratchLocalQuickFixDocument(t, "Option Explicit\nSub Work()\n"+body+"\nEnd Sub\n")
			fixes, err := linter.QuickFixesParsedContext(t.Context(), doc)
			if err != nil {
				t.Fatal(err)
			}
			for _, fix := range fixes {
				if fix.Diagnostic.Code == "VB020" {
					t.Fatalf("unsafe fix = %+v", fix)
				}
			}
		})
	}
}

func TestUnusedLocalQuickFixProcedureOwnership(t *testing.T) {
	for name, source := range map[string]string{
		"module":                "Option Explicit\nDim scratch As Long\nSub Work()\nEnd Sub\n",
		"static procedure":      "Option Explicit\nStatic Sub Work()\nDim scratch As Long\nEnd Sub\n",
		"parameter conflict":    "Option Explicit\nSub Work(scratch As Long)\nDim scratch As Long\nEnd Sub\n",
		"conditional procedure": "Option Explicit\n#If VBA7 Then\nSub Work()\nDim scratch As Long\nEnd Sub\n#End If\n",
	} {
		t.Run(name, func(t *testing.T) {
			linter, doc := scratchLocalQuickFixDocument(t, source)
			fixes, err := linter.QuickFixesParsedContext(t.Context(), doc)
			if err != nil || len(fixes) != 0 {
				t.Fatalf("fixes = %+v, error = %v", fixes, err)
			}
		})
	}
	linter, doc := scratchLocalQuickFixDocument(t, "Option Explicit\nProperty Get Value() As Long\nDim scratch As Long\nEnd Property\nProperty Let Value(ByVal rhs As Long)\nDim scratch As Long\nscratch = rhs\nEnd Property\n")
	fixes, err := linter.QuickFixesParsedContext(t.Context(), doc)
	if err != nil || len(fixes) != 1 || fixes[0].Diagnostic.Line != 3 {
		t.Fatalf("property accessor fixes = %+v, error = %v", fixes, err)
	}
}

func scratchLocalQuickFixDocument(t *testing.T, source string) (Linter, *vbaast.ParsedDocument) {
	t.Helper()
	doc, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(doc.Close)
	return Linter{RootDir: ".", Config: config.Default()}, doc
}
