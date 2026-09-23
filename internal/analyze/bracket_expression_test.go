package analyze

import (
	"path/filepath"
	"testing"

	vbaast "github.com/harumiWeb/xlflow/internal/vba/ast"
	"github.com/harumiWeb/xlflow/internal/vba/procedureir"
)

func TestBracketExpressionFindingsReportsBareHostExpressions(t *testing.T) {
	const source = `Option Explicit
Public Sub Run()
  result = [A1]
  result = Foo([A2])
  result = Me.[Member]
  result = values(index)
  Dim [Declared] As String
  result = "literal [A4]"
  ' [A5]
End Sub
`
	root := t.TempDir()
	ir, lines := buildBracketExpressionTestIR(t, source)
	file := parsedFile{Path: filepath.Join(root, "Main.bas"), Lines: lines, Module: "Main"}
	procIR := &ir.Procedures[0]
	proc := sourceProcedure{IR: procIR, Name: procIR.Symbol.Name, EndLine: procIR.Symbol.DeclarationRange.EndLine}

	findings := bracketExpressionFindings(root, file, proc)
	if len(findings) != 2 {
		t.Fatalf("VBA262 findings = %+v, want two bare bracket expressions", findings)
	}
	if findings[0].Line != 3 || findings[0].Column < 1 || findings[0].File != "Main.bas" {
		t.Fatalf("first VBA262 finding = %+v, want [A1] location in Main.bas", findings[0])
	}
	if findings[1].Line != 4 || findings[1].ScopeEndLine != proc.EndLine {
		t.Fatalf("second VBA262 finding = %+v, want Foo([A2]) location scoped to procedure", findings[1])
	}
}

func TestBracketExpressionCandidatesExcludeMembersAndDeclarations(t *testing.T) {
	const source = `Public Sub Run()
  result = Me.[Member]
  Dim [Declared] As String
  result = values(index)
  result = "[Literal]"
  ' [Comment]
End Sub
`
	ir, _ := buildBracketExpressionTestIR(t, source)
	if got := bracketExpressionCandidates(&ir.Procedures[0]); len(got) != 0 {
		t.Fatalf("bracket expression candidates = %+v, want none", got)
	}
}

func buildBracketExpressionTestIR(t *testing.T, source string) (procedureir.DocumentIR, []string) {
	t.Helper()
	document, err := vbaast.ParseDocument("Main.bas", []byte(source))
	if err != nil {
		t.Fatal(err)
	}
	defer document.Close()
	ir, err := procedureir.BuildParsedContext(t.Context(), procedureir.BuildOptions{
		Path:       "Main.bas",
		ModuleName: "Main",
		ModuleKind: "standard",
	}, document)
	if err != nil {
		t.Fatal(err)
	}
	if len(ir.Procedures) != 1 {
		t.Fatalf("procedure count = %d, want one", len(ir.Procedures))
	}
	return ir, normalizedSourceLines(source)
}
